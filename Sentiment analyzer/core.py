"""
core.py
-------
Contains the core business logic, including database connections, 
data visualization, markdown parsing, and the LLM generation pipeline.
"""

import os
import io
import logging
import pandas as pd
import chromadb
import markdown
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from ollama import Client 
from chromadb.utils import embedding_functions

# Import configurations directly from config.py
from config import OLLAMA_HOST, LLM_MODEL, EMBED_MODEL, DB_FILE, COMPANY_NAME, REPORT_STRUCTURE

logger = logging.getLogger(__name__)

# --- PARSING & VISUALIZATION ---

class MarkdownParser:
    """Translates the LLM's raw Markdown output into richly formatted Microsoft Word elements."""
    @staticmethod
    def add_markdown_to_doc(doc, md_text):
        html = markdown.markdown(md_text)
        soup = BeautifulSoup(html, 'html.parser')
        
        for element in soup.children:
            if element.name is None: continue
            
            if element.name in ['h1', 'h2', 'h3']:
                level = int(element.name[1]) + 1 
                doc.add_heading(element.get_text().strip(), level=level)
            elif element.name == 'p':
                text = element.get_text().strip()
                if not text: continue
                p = doc.add_paragraph()
                MarkdownParser._process_inline(p, element)
            elif element.name in ['ul', 'ol']:
                style = 'List Bullet' if element.name == 'ul' else 'List Number'
                for li in element.find_all('li'):
                    p = doc.add_paragraph(style=style)
                    MarkdownParser._process_inline(p, li)

    @staticmethod
    def _process_inline(paragraph, element):
        for child in element.children:
            if child.name in ['strong', 'b']:
                paragraph.add_run(child.get_text()).bold = True
            elif child.name in ['em', 'i']:
                paragraph.add_run(child.get_text()).italic = True
            elif child.name is None:
                text = str(child).replace('\n', ' ')
                paragraph.add_run(text)
            else:
                MarkdownParser._process_inline(paragraph, child)

class ChartEngine:
    """Handles generation of all visual data elements using Matplotlib."""
    @staticmethod
    def create_sentiment_chart(csat_score):
        try:
            score = float(csat_score.split('/')[0])
            max_score = float(csat_score.split('/')[1]) if '/' in csat_score else 5.0
            
            plt.figure(figsize=(5, 1.5))
            plt.barh(['Kepuasan Klien'], [score], color='#2563eb', alpha=0.8)
            plt.xlim(0, max_score)
            plt.title(f"Tingkat Kepuasan (CSAT): {csat_score}", fontsize=10, fontweight='bold')
            plt.gca().spines['top'].set_visible(False)
            plt.gca().spines['right'].set_visible(False)
            
            img = io.BytesIO()
            plt.savefig(img, format='png', bbox_inches='tight', dpi=150)
            plt.close()
            img.seek(0)
            return img
        except Exception as e:
            logger.error(f"Failed to create sentiment chart: {e}")
            return None

    @staticmethod
    def create_nps_distribution_chart(df):
        try:
            nps_scores = pd.to_numeric(df['NPS'], errors='coerce').dropna()
            promoters = len(nps_scores[nps_scores >= 9])
            passives = len(nps_scores[(nps_scores >= 7) & (nps_scores <= 8)])
            detractors = len(nps_scores[nps_scores <= 6])
            
            sizes = [promoters, passives, detractors]
            labels = ['Promoters (9-10)', 'Passives (7-8)', 'Detractors (0-6)']
            colors = ['#10b981', '#fbbf24', '#ef4444'] 
            
            if sum(sizes) == 0: return None
            
            plt.figure(figsize=(4, 4))
            plt.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=140)
            plt.title("Distribusi Keseluruhan NPS", fontsize=11, fontweight='bold')
            
            img = io.BytesIO()
            plt.savefig(img, format='png', bbox_inches='tight', dpi=150)
            plt.close()
            img.seek(0)
            return img
        except Exception as e: 
            logger.error(f"Failed to create NPS distribution chart: {e}")
            return None

# --- DATABASE & GENERATION ---

class FeedbackDatabase:
    """Handles reading the CSV file and syncing it with the local ChromaDB vector store."""
    def __init__(self):
        logger.info("Initializing Database connection...")
        self.chroma = chromadb.Client()
        self.embed_fn = embedding_functions.OllamaEmbeddingFunction(
            url=f"{OLLAMA_HOST}/api/embeddings", 
            model_name=EMBED_MODEL
        )
        self.collection = self.chroma.get_or_create_collection(
            name="feedback_db", 
            embedding_function=self.embed_fn
        )
        self.df = None
        self.refresh_data()

    def refresh_data(self):
        if not os.path.exists(DB_FILE):
            logger.warning(f"Database file {DB_FILE} not found.")
            return False
            
        try:
            self.df = pd.read_csv(DB_FILE)
            existing_ids = self.collection.get()['ids']
            if existing_ids: 
                self.collection.delete(existing_ids)
            
            ids, docs, metas = [], [], []
            for idx, row in self.df.iterrows():
                client = row.get('Client/Partner', 'Unknown')
                feedback = row.get('Raw Feedback Text', '')
                text_rep = f"Client: {client} | Feedback: {feedback}"
                
                ids.append(str(idx))
                docs.append(text_rep)
                metas.append(row.to_dict())
                
            if ids: 
                self.collection.add(documents=docs, metadatas=metas, ids=ids)
                
            logger.info(f"Successfully loaded {len(ids)} rows into the database.")
            return True
        except Exception as e: 
            logger.error(f"Failed to refresh database: {e}", exc_info=True)
            return False

    def get_client_feedback(self, client_name):
        if self.df is None: return ""
        client_data = self.df[self.df['Client/Partner'] == client_name]
        if client_data.empty: return "No data found."
        
        return "\n".join([
            f"Date: {row.get('Survey Date')} | CSAT: {row.get('CSAT Score')} | Feedback: {row.get('Raw Feedback Text')}"
            for _, row in client_data.iterrows()
        ])
        
    def get_overall_feedback(self):
        if self.df is None or self.df.empty: return ""
        return "\n".join([
            f"Client: {row['Client/Partner']} | CSAT: {row['CSAT Score']} | Feedback: {row['Raw Feedback Text']}"
            for _, row in self.df.iterrows()
        ])

class ReportGenerator:
    """Orchestrates fetching data, prompting the LLM, and building the Docx file."""
    def __init__(self, db_instance):
        self.ollama = Client(host=OLLAMA_HOST)
        self.db = db_instance

    def generate(self, client_name):
        logger.info(f"Starting report generation for: {client_name}")
        doc = Document()
        
        style = doc.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)
        
        is_overall = (client_name == "ALL_OVERALL")
        
        if is_overall:
            raw_feedback = self.db.get_overall_feedback()
            report_title = "Laporan Analisis Sentimen Makro\n(Seluruh Klien)"
            filename = "Insight_Report_OVERALL"
        else:
            raw_feedback = self.db.get_client_feedback(client_name)
            report_title = f"Laporan Analisis Sentimen & Feedback\nKlien: {client_name}"
            filename = f"Insight_Report_{client_name.replace(' ', '_')}"
            
        doc.add_heading(report_title, level=0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Disusun secara otomatis untuk Manajemen {COMPANY_NAME}").alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_page_break()
        
        for chap in REPORT_STRUCTURE:
            logger.info(f"Generating section: {chap['title']}")
            
            prompt = f"""
            You are an expert Data Analyst and Consultant working for {COMPANY_NAME}.
            You are analyzing the following feedback dataset:
            RAW DATA: {raw_feedback}
            TASK: {chap['instructions']}
            RULES:
            - Write the content entirely in professional Indonesian.
            - Keep it deeply analytical and actionable.
            - DO NOT include the chapter title in your output.
            """
            
            try:
                response = self.ollama.chat(
                    model=LLM_MODEL, 
                    messages=[{'role':'user', 'content': prompt}]
                )
                content = response['message']['content']
                
                doc.add_heading(chap['title'], level=1)
                
                if chap.get("visual_intent") == "sentiment_chart":
                    if is_overall:
                        chart_img = ChartEngine.create_nps_distribution_chart(self.db.df)
                        if chart_img: doc.add_picture(chart_img, width=Inches(3.5))
                    else:
                        client_row = self.db.df[self.db.df['Client/Partner'] == client_name].iloc[0]
                        csat = str(client_row.get('CSAT Score', '0/5'))
                        chart_img = ChartEngine.create_sentiment_chart(csat)
                        if chart_img: doc.add_picture(chart_img, width=Inches(4))
                
                MarkdownParser.add_markdown_to_doc(doc, content)
                doc.add_paragraph() 
                
            except Exception as e:
                logger.error(f"Error generating chapter '{chap['title']}': {e}", exc_info=True)
                doc.add_paragraph("[Terjadi kesalahan sistem saat membuat bagian ini.]")

        logger.info("Report generation complete.")
        return doc, filename