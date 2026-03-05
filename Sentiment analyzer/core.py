"""
core.py
-------
Handles database ingestion, data visualization, markdown parsing, 
and the LLM pipeline for generating the Word document.
Includes advanced docx formatting, native table rendering, 
recursive sub-bullet handling, and strict emoji sanitization.
"""

import os
import io
import re
import logging
import datetime
import pandas as pd
import chromadb
import markdown
import matplotlib
import matplotlib.dates as mdates
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn
from ollama import Client 
from chromadb.utils import embedding_functions

from config import OLLAMA_URL, AI_MODEL, EMBEDDING_MODEL, DATA_SOURCE, ORG_NAME, REPORT_SECTIONS

logger = logging.getLogger(__name__)

class WordFormatter:
    """Helper class to convert raw Markdown into highly formatted Word document styling."""
    
    @staticmethod
    def apply_markdown(document, md_text):
        md_text = re.sub(r'[\U00010000-\U0010ffff]', '', md_text)
        md_text = re.sub(r'^(#{1,6})([^ #\n])', r'\1 \2', md_text, flags=re.MULTILINE)
        
        html_content = markdown.markdown(md_text, extensions=['tables'])
        parsed_html = BeautifulSoup(html_content, 'html.parser')
        
        for tag in parsed_html.children:
            if tag.name is None: 
                continue
            
            if tag.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
                heading_level = min(int(tag.name[1]) + 1, 9) 
                clean_text = re.sub(r'^#+\s*', '', tag.get_text().strip())
                heading = document.add_heading(clean_text, level=heading_level)
                heading.paragraph_format.space_before = Pt(18)
                heading.paragraph_format.space_after = Pt(6)
                heading.paragraph_format.keep_with_next = True
                
            elif tag.name == 'p':
                text_content = tag.get_text().strip()
                if not text_content: 
                    continue
                
                rogue_heading_match = re.match(r'^(#{1,6})\s*(.*)', text_content)
                if rogue_heading_match:
                    hashes = rogue_heading_match.group(1)
                    clean_text = rogue_heading_match.group(2).strip()
                    heading_level = min(len(hashes) + 1, 9)
                    heading = document.add_heading(clean_text, level=heading_level)
                    heading.paragraph_format.space_before = Pt(18)
                    heading.paragraph_format.space_after = Pt(6)
                    heading.paragraph_format.keep_with_next = True
                    continue 
                    
                paragraph = document.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                paragraph.paragraph_format.space_after = Pt(10)
                WordFormatter._style_inline_text(paragraph, tag)
                
            elif tag.name in ['ul', 'ol']:
                # Pass the list to our new recursive processor
                WordFormatter._process_list(document, tag, level=0)
                    
            elif tag.name == 'table':
                first_row = tag.find('tr')
                if not first_row:
                    continue
                num_cols = len(first_row.find_all(['th', 'td']))
                
                word_table = document.add_table(rows=0, cols=num_cols)
                word_table.style = 'Table Grid'
                word_table.autofit = True
                
                for tr in tag.find_all('tr'):
                    row_cells = word_table.add_row().cells
                    col_idx = 0
                    for cell in tr.find_all(['th', 'td']):
                        if col_idx < num_cols:
                            cell_paragraph = row_cells[col_idx].paragraphs[0]
                            WordFormatter._style_inline_text(cell_paragraph, cell)
                            
                            if cell.name == 'th':
                                cell_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                for run in cell_paragraph.runs:
                                    run.bold = True
                        col_idx += 1
                
                document.add_paragraph().paragraph_format.space_after = Pt(12)

    @staticmethod
    def _process_list(document, list_tag, level=0):
        """Recursively parses nested lists and applies dynamic indentation/styles."""
        is_ul = list_tag.name == 'ul'
        base_style = 'List Bullet' if is_ul else 'List Number'
        
        # Word supports native nested list styles like 'List Bullet 2', 'List Bullet 3'
        style_name = base_style if level == 0 else f'{base_style} {min(level + 1, 5)}'
        
        for li in list_tag.children:
            if li.name != 'li': 
                continue
                
            try:
                paragraph = document.add_paragraph(style=style_name)
            except KeyError:
                # Fallback if the default Word template is missing the nested style
                paragraph = document.add_paragraph(style=base_style)
            
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            # Push the text further right for each level of nesting
            paragraph.paragraph_format.left_indent = Inches(0.25 + (level * 0.35))
            paragraph.paragraph_format.space_after = Pt(4)
            
            for child in li.children:
                if child.name in ['ul', 'ol']:
                    # Sub-bullet found! Recurse deeper.
                    WordFormatter._process_list(document, child, level + 1)
                else:
                    WordFormatter._style_single_node(paragraph, child)

    @staticmethod
    def _style_inline_text(paragraph, html_element):
        for child in html_element.children:
            if child.name in ['ul', 'ol']: 
                continue # Nested lists inside paragraphs are skipped here to prevent duplication
            WordFormatter._style_single_node(paragraph, child)

    @staticmethod
    def _style_single_node(paragraph, child):
        if child.name in ['strong', 'b']:
            clean_text = child.get_text().replace('*', '').strip()
            run = paragraph.add_run(clean_text)
            run.bold = True
            if child.next_sibling and isinstance(child.next_sibling, str) and not child.next_sibling.startswith((' ', '.', ',')):
                paragraph.add_run(' ')
        elif child.name in ['em', 'i']:
            clean_text = child.get_text().replace('_', '').strip()
            run = paragraph.add_run(clean_text)
            run.italic = True
        elif child.name is None:
            clean_text = str(child).replace('\n', ' ').replace('**', '').replace('__', '')
            paragraph.add_run(clean_text)
        else:
            for subchild in child.children:
                WordFormatter._style_single_node(paragraph, subchild)


class GraphMaker:
    """Creates visual representations of the survey data."""
    
    @staticmethod
    def build_nps_pie_chart(dataframe):
        try:
            scores = pd.to_numeric(dataframe['NPS'], errors='coerce').dropna()
            
            promoters = len(scores[scores >= 9])
            passives = len(scores[(scores >= 7) & (scores <= 8)])
            detractors = len(scores[scores <= 6])
            
            categories = [promoters, passives, detractors]
            labels = ['Promoters (9-10)', 'Passives (7-8)', 'Detractors (0-6)']
            colors = ['#10b981', '#fbbf24', '#ef4444'] 
            
            if sum(categories) == 0: 
                return None
            
            plt.figure(figsize=(5, 5))
            plt.pie(categories, labels=labels, colors=colors, autopct='%1.1f%%', startangle=140, textprops={'fontsize': 10})
            plt.title("Distribusi Skor NPS", fontsize=12, fontweight='bold', pad=15)
            
            image_stream = io.BytesIO()
            plt.savefig(image_stream, format='png', bbox_inches='tight', dpi=200)
            plt.close()
            image_stream.seek(0)
            return image_stream
            
        except Exception as err: 
            logger.error(f"Failed to build NPS pie chart: {err}")
            return None

    @staticmethod
    def build_trend_line_chart(dataframe):
        try:
            df = dataframe.copy()
            df['CSAT_Val'] = df['CSAT Score'].astype(str).str.split('/').str[0].astype(float)
            df['Survey Date'] = pd.to_datetime(df['Survey Date'])
            
            trend_data = df.groupby('Survey Date')['CSAT_Val'].mean().reset_index()
            trend_data = trend_data.sort_values('Survey Date')
            
            if trend_data.empty: return None

            plt.figure(figsize=(7, 3.5))
            plt.plot(trend_data['Survey Date'], trend_data['CSAT_Val'], marker='o', linestyle='-', color='#1e3a8a', linewidth=2.5)
            
            plt.ylim(0, 5.5)
            plt.title("Tren Kepuasan (CSAT) Seiring Waktu", fontsize=12, fontweight='bold', pad=15)
            plt.ylabel("Rata-rata CSAT (Skala 5)", fontsize=10)
            plt.grid(True, linestyle='--', alpha=0.5)
            
            plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%d %b %Y'))
            plt.gcf().autofmt_xdate() 
            
            plt.gca().spines['top'].set_visible(False)
            plt.gca().spines['right'].set_visible(False)
            
            image_stream = io.BytesIO()
            plt.savefig(image_stream, format='png', bbox_inches='tight', dpi=200)
            plt.close()
            image_stream.seek(0)
            return image_stream
            
        except Exception as err:
            logger.error(f"Failed to build trend chart: {err}")
            return None


class SurveyData:
    """Manages the CSV data and ChromaDB vector store synchronization."""
    
    def __init__(self):
        logger.info("Connecting to local database...")
        self.chroma_client = chromadb.Client()
        self.embedder = embedding_functions.OllamaEmbeddingFunction(
            url=f"{OLLAMA_URL}/api/embeddings", 
            model_name=EMBEDDING_MODEL
        )
        self.db_collection = self.chroma_client.get_or_create_collection(
            name="survey_records", 
            embedding_function=self.embedder
        )
        self.raw_data = None
        self.active_filtered_data = None
        self.load_csv()

    def load_csv(self):
        if not os.path.exists(DATA_SOURCE):
            logger.warning(f"Could not find data file at {DATA_SOURCE}.")
            return False
            
        try:
            self.raw_data = pd.read_csv(DATA_SOURCE)
            self.raw_data['Survey Date'] = pd.to_datetime(self.raw_data['Survey Date'])
            
            existing_records = self.db_collection.get()['ids']
            if existing_records: 
                self.db_collection.delete(existing_records)
            
            record_ids, documents, metadata = [], [], []
            
            for index, row in self.raw_data.iterrows():
                partner = row.get('Client/Partner', 'Unknown')
                text_content = row.get('Raw Feedback Text', '')
                document_string = f"Client: {partner} | Feedback: {text_content}"
                
                record_ids.append(str(index))
                documents.append(document_string)
                metadata.append({k: str(v) for k, v in row.to_dict().items()})
                
            if record_ids: 
                self.db_collection.add(documents=documents, metadatas=metadata, ids=record_ids)
                
            logger.info(f"Loaded {len(record_ids)} records into ChromaDB.")
            return True
            
        except Exception as err: 
            logger.error(f"Database ingestion failed: {err}", exc_info=True)
            return False

    def filter_by_timeframe(self, duration_code):
        if self.raw_data is None or self.raw_data.empty: 
            return "DATABASE KOSONG."
        
        latest_record_date = self.raw_data['Survey Date'].max()
        
        time_offsets = {
            '1w': pd.Timedelta(weeks=1),
            '1m': pd.DateOffset(months=1),
            '3m': pd.DateOffset(months=3),
            '6m': pd.DateOffset(months=6),
            '1y': pd.DateOffset(years=1)
        }
        
        start_date = latest_record_date - time_offsets.get(duration_code, pd.DateOffset(years=10))
        
        mask = (self.raw_data['Survey Date'] >= start_date) & (self.raw_data['Survey Date'] <= latest_record_date)
        filtered_df = self.raw_data.loc[mask]
        
        self.active_filtered_data = filtered_df 
        
        if filtered_df.empty: 
            return "TIDAK ADA DATA PADA PERIODE INI."
        
        compiled_text = []
        for _, row in filtered_df.iterrows():
            date_str = row['Survey Date'].strftime('%Y-%m-%d')
            compiled_text.append(
                f"Date: {date_str} | Source: {row.get('Feedback Source')} | "
                f"CSAT: {row.get('CSAT Score')} | Text: {row.get('Raw Feedback Text')}"
            )
            
        return "\n".join(compiled_text)


class HRReportBuilder:
    """Coordinates data fetching, LLM generation, and Word doc compilation."""
    
    def __init__(self, database):
        self.llm_client = Client(host=OLLAMA_URL)
        self.database = database

    def _setup_document_styles(self, document):
        """Configures the foundational typography for the report."""
        normal_style = document.styles['Normal']
        normal_style.font.name = 'Arial'
        normal_style.font.size = Pt(11)
        normal_style.font.color.rgb = RGBColor(30, 41, 59) 
        
        for i in range(1, 5):
            try:
                heading_style = document.styles[f'Heading {i}']
                heading_style.font.name = 'Arial'
                heading_style.font.bold = True
                heading_style.font.color.rgb = RGBColor(15, 23, 42)
            except KeyError:
                pass 

    def compile_report(self, duration_code, duration_label):
        logger.info(f"Generating HR insight report for timeframe: {duration_code}")
        document = Document()
        self._setup_document_styles(document)
        
        dataset_context = self.database.filter_by_timeframe(duration_code)
        
        document.add_paragraph() 
        document.add_paragraph()
        
        title_para = document.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_para.add_run("LAPORAN AUDIT SENTIMEN & KELUHAN")
        title_run.font.size = Pt(22)
        title_run.bold = True
        
        subtitle_para = document.add_paragraph()
        subtitle_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        subtitle_run = subtitle_para.add_run(f"Periode Analisis: {duration_label}")
        subtitle_run.font.size = Pt(14)
        subtitle_run.italic = True
        
        document.add_paragraph()
        meta_para = document.add_paragraph()
        meta_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        meta_run = meta_para.add_run(f"Disusun secara otomatis untuk HR & Manajemen Eksekutif\n{ORG_NAME}\nTanggal Terbit: {datetime.datetime.now().strftime('%d %B %Y')}")
        meta_run.font.size = Pt(11)
        meta_run.font.color.rgb = RGBColor(100, 116, 139)

        document.add_page_break()
        
        file_name = f"HR_Audit_Report_{duration_code}"
        
        for section in REPORT_SECTIONS:
            logger.info(f"Writing chapter: {section['title']}")
            
            prompt_context = f"""
            You are a highly critical Data Analyst and HR Diagnostician for {ORG_NAME}.
            Analyze the following feedback dataset collected over the period: {duration_label}.
            
            RAW DATA: 
            {dataset_context}
            
            TASK: {section['instructions']}
            """
            
            try:
                response = self.llm_client.chat(
                    model=AI_MODEL, 
                    messages=[{'role':'user', 'content': prompt_context}],
                    options={'temperature': 0.1} 
                )
                generated_text = response['message']['content']
                
                heading = document.add_heading(section['title'], level=1)
                heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                if section.get("include_trend_chart") and self.database.active_filtered_data is not None:
                    
                    trend_chart = GraphMaker.build_trend_line_chart(self.database.active_filtered_data)
                    if trend_chart: 
                        para = document.add_paragraph()
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        para.add_run().add_picture(trend_chart, width=Inches(6.0))
                        
                    pie_chart = GraphMaker.build_nps_pie_chart(self.database.active_filtered_data)
                    if pie_chart: 
                        para = document.add_paragraph()
                        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        para.add_run().add_picture(pie_chart, width=Inches(4.5))
                
                WordFormatter.apply_markdown(document, generated_text)
                
                if section != REPORT_SECTIONS[-1]:
                    document.add_page_break()
                
            except Exception as err:
                logger.error(f"Failed writing section '{section['title']}': {err}", exc_info=True)
                document.add_paragraph("[Sistem gagal menghasilkan analisis untuk bagian ini.]")

        logger.info("Report assembly finished successfully.")
        return document, file_name