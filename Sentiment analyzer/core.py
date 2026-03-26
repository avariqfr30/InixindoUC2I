import concurrent.futures
import io
import logging
import os
import re
import textwrap
from datetime import datetime
from urllib.parse import urlparse

import chromadb
import markdown
import matplotlib
import matplotlib.patches as patches
import matplotlib.pyplot as plt
import pandas as pd
import requests
from bs4 import BeautifulSoup, NavigableString, Tag
from chromadb.config import Settings
from chromadb.utils import embedding_functions
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from ollama import Client
from sqlalchemy import create_engine

from config import (
    CX_ANALYSIS_SYSTEM_PROMPT,
    CX_SENTIMENT_STRUCTURE,
    CSV_PATH,
    DATA_DIR,
    DEFAULT_COLOR,
    EMBED_MODEL,
    LLM_MODEL,
    OLLAMA_HOST,
    OSINT_BASE_QUERIES,
    OSINT_MAX_SIGNALS,
    OSINT_RECENCY,
    OSINT_RESULTS_PER_QUERY,
    OSINT_SEARCH_LANGUAGE,
    OSINT_SEARCH_REGION,
    OSINT_TOPIC_QUERY_TEMPLATE,
    PERSONAS,
    SERPER_API_KEY,
    WRITER_FIRM_NAME,
)

matplotlib.use("Agg")

logger = logging.getLogger(__name__)


def append_field(paragraph, instruction):
    run = paragraph.add_run()

    field_begin = OxmlElement("w:fldChar")
    field_begin.set(qn("w:fldCharType"), "begin")

    field_instruction = OxmlElement("w:instrText")
    field_instruction.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    field_instruction.text = instruction

    field_separator = OxmlElement("w:fldChar")
    field_separator.set(qn("w:fldCharType"), "separate")

    field_end = OxmlElement("w:fldChar")
    field_end.set(qn("w:fldCharType"), "end")

    run._r.extend([field_begin, field_instruction, field_separator, field_end])


class KnowledgeBase:
    def __init__(self, db_uri):
        os.makedirs(DATA_DIR, exist_ok=True)
        self.engine = create_engine(db_uri)
        self.chroma = chromadb.Client(Settings(anonymized_telemetry=False))
        self.embed_fn = embedding_functions.OllamaEmbeddingFunction(
            url=f"{OLLAMA_HOST}/api/embeddings",
            model_name=EMBED_MODEL,
        )
        self.collection = self.chroma.get_or_create_collection(
            name="cx_holistic_db",
            embedding_function=self.embed_fn,
        )
        self.df = None
        self.refresh_data()

    def refresh_data(self):
        try:
            self.df = pd.read_sql("SELECT * FROM feedback", self.engine)
        except Exception:
            if not os.path.exists(CSV_PATH):
                logger.error("Gagal memuat data: file %s tidak ditemukan.", CSV_PATH)
                return False

            raw_df = pd.read_csv(CSV_PATH)
            raw_df.columns = [column.strip() for column in raw_df.columns]
            raw_df.to_sql("feedback", self.engine, index=False, if_exists="replace")
            self.df = raw_df

        try:
            existing_ids = self.collection.get().get("ids", [])
            if existing_ids:
                self.collection.delete(existing_ids)

            ids, documents, metadata = [], [], []
            for index, row in self.df.iterrows():
                text_representation = " | ".join(
                    f"{column}: {value}" for column, value in row.items()
                )
                ids.append(str(index))
                documents.append(text_representation)
                metadata.append(row.astype(str).to_dict())

            if ids:
                logger.info(
                    "Mengirim %s feedback ke endpoint embedding Ollama (%s).",
                    len(ids),
                    OLLAMA_HOST,
                )
                self.collection.add(documents=documents, metadatas=metadata, ids=ids)
        except Exception as exc:
            logger.error("Gagal memperbarui vector store: %s", exc)
            return False

        return True

    def query(self, timeframe, context_keywords=""):
        query_text = (
            f"General feedback, complaints, praise, and operational issues. "
            f"{context_keywords}"
        ).strip()
        query_payload = {"query_texts": [query_text], "n_results": 25}
        if timeframe:
            query_payload["where"] = {"Rentang Waktu": timeframe}

        try:
            result = self.collection.query(**query_payload)
            documents = result.get("documents", [[]])
            if documents and documents[0]:
                return "\n---\n".join(documents[0])
        except Exception as exc:
            logger.error("Query error: %s", exc)

        return "Tidak ada data feedback internal untuk periode ini."


class Researcher:
    SERPER_ENDPOINT = "https://google.serper.dev/search"
    INVALID_API_KEYS = {
        "",
        "YOUR_SERPER_API_KEY",
        "masukkan_api_key_serper_anda_disini",
    }

    @staticmethod
    def _is_enabled():
        return (SERPER_API_KEY or "").strip() not in Researcher.INVALID_API_KEYS

    @staticmethod
    def _search_serper(query, max_results=OSINT_RESULTS_PER_QUERY):
        payload = {
            "q": query,
            "num": max_results,
            "gl": OSINT_SEARCH_REGION,
            "hl": OSINT_SEARCH_LANGUAGE,
        }
        if OSINT_RECENCY:
            payload["tbs"] = OSINT_RECENCY

        headers = {
            "X-API-KEY": SERPER_API_KEY,
            "Content-Type": "application/json",
        }

        response = requests.post(
            Researcher.SERPER_ENDPOINT,
            headers=headers,
            json=payload,
            timeout=10,
        )
        response.raise_for_status()
        return response.json()

    @staticmethod
    def _extract_items(query, payload):
        items = []

        for position, entry in enumerate(payload.get("organic", []), start=1):
            items.append(
                {
                    "query": query,
                    "title": entry.get("title", "Tanpa Judul"),
                    "snippet": entry.get("snippet", "").strip(),
                    "url": entry.get("link", "").strip(),
                    "date": entry.get("date", "").strip(),
                    "source_type": "organic",
                    "position": position,
                }
            )

        for position, entry in enumerate(payload.get("news", []), start=1):
            items.append(
                {
                    "query": query,
                    "title": entry.get("title", "Tanpa Judul"),
                    "snippet": entry.get("snippet", "").strip(),
                    "url": entry.get("link", "").strip(),
                    "date": entry.get("date", "").strip(),
                    "source_type": "news",
                    "position": position,
                }
            )

        return [item for item in items if item["url"]]

    @staticmethod
    def _deduplicate_items(items):
        seen_keys = set()
        unique_items = []

        for item in items:
            key = item["url"] or f"{item['title']}::{item['snippet']}"
            if key in seen_keys:
                continue
            seen_keys.add(key)
            unique_items.append(item)

        return unique_items

    @staticmethod
    def _score_items(items, context_text):
        keywords = {
            token.lower()
            for token in re.findall(r"[A-Za-z]{4,}", context_text)
            if token.lower() not in {"yang", "dengan", "untuk", "pada", "from", "this"}
        }

        for item in items:
            corpus = f"{item['title']} {item['snippet']}".lower()
            coverage_score = sum(1 for keyword in keywords if keyword in corpus)
            freshness_bonus = 1 if item["source_type"] == "news" else 0
            ranking_penalty = (item["position"] - 1) * 0.05
            item["score"] = coverage_score + freshness_bonus - ranking_penalty

        return sorted(items, key=lambda value: value["score"], reverse=True)

    @staticmethod
    def _source_domain(url):
        try:
            netloc = urlparse(url).netloc.lower()
            return netloc.replace("www.", "") if netloc else "unknown"
        except Exception:
            return "unknown"

    @staticmethod
    def _format_osint_brief(items, title):
        if not items:
            return "Tidak ada sinyal OSINT eksternal yang dapat digunakan untuk benchmark periode ini."

        lines = [f"{title}:"]
        for index, item in enumerate(items, start=1):
            date_part = f" | tanggal={item['date']}" if item["date"] else ""
            lines.append(
                (
                    f"{index}. {item['title']} | {item['snippet']} "
                    f"| sumber={Researcher._source_domain(item['url'])}"
                    f"{date_part} | url={item['url']}"
                ).strip()
            )

        return "\n".join(lines)

    @staticmethod
    def _run_query_batch(queries, max_signals=OSINT_MAX_SIGNALS):
        collected = []

        with concurrent.futures.ThreadPoolExecutor(
            max_workers=min(4, max(1, len(queries)))
        ) as pool:
            future_map = {
                pool.submit(Researcher._search_serper, query): query for query in queries
            }
            for future in concurrent.futures.as_completed(future_map):
                query = future_map[future]
                try:
                    payload = future.result()
                    collected.extend(Researcher._extract_items(query, payload))
                except Exception as exc:
                    logger.warning("OSINT query gagal (%s): %s", query, exc)

        deduplicated = Researcher._deduplicate_items(collected)
        ranked = Researcher._score_items(deduplicated, " ".join(queries))
        return ranked[:max_signals]

    @staticmethod
    def get_macro_trends(timeframe, notes=""):
        if not Researcher._is_enabled():
            return "Data OSINT eksternal tidak tersedia (SERPER_API_KEY belum diatur)."

        scope = timeframe or "periode terbaru"
        compact_notes = re.sub(r"\s+", " ", notes).strip()
        contextual_query = (
            f"benchmark sentimen pelanggan pelatihan dan konsultasi IT Indonesia {scope}"
        )
        if compact_notes:
            contextual_query += f" {compact_notes[:140]}"

        queries = [
            f"{query} {scope}" for query in OSINT_BASE_QUERIES
        ] + [contextual_query]

        findings = Researcher._run_query_batch(queries)
        return Researcher._format_osint_brief(
            findings,
            "Sinyal OSINT Makro (Indonesia)",
        )

    @staticmethod
    def get_topic_trends(timeframe, focus_keywords, notes=""):
        if not Researcher._is_enabled():
            return "Data OSINT topikal tidak tersedia (SERPER_API_KEY belum diatur)."

        query = OSINT_TOPIC_QUERY_TEMPLATE.format(
            timeframe=timeframe or "periode terbaru",
            focus_keywords=focus_keywords,
            notes=re.sub(r"\s+", " ", notes).strip()[:120],
        )

        findings = Researcher._run_query_batch([query], max_signals=4)
        return Researcher._format_osint_brief(
            findings,
            "Sinyal OSINT Fokus Bab",
        )


class StyleEngine:
    @staticmethod
    def _configure_text_style(style, font_name, font_size, color, bold=False):
        style.font.name = font_name
        style.font.size = Pt(font_size)
        style.font.color.rgb = RGBColor(*color)
        style.font.bold = bold

    @staticmethod
    def apply_document_styles(doc, theme_color):
        for section in doc.sections:
            section.top_margin = Cm(2.54)
            section.bottom_margin = Cm(2.54)
            section.left_margin = Cm(2.54)
            section.right_margin = Cm(2.54)

            footer = section.footer
            footer_paragraph = (
                footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            )
            footer_paragraph.clear()
            footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            footer_run = footer_paragraph.add_run(
                "STRICTLY CONFIDENTIAL | Inixindo Jogja Executive Report | Page "
            )
            footer_run.font.name = "Calibri"
            footer_run.font.size = Pt(9)
            footer_run.font.color.rgb = RGBColor(128, 128, 128)
            append_field(footer_paragraph, "PAGE")

        normal_style = doc.styles["Normal"]
        StyleEngine._configure_text_style(
            normal_style,
            font_name="Calibri",
            font_size=11,
            color=(33, 37, 41),
        )
        normal_format = normal_style.paragraph_format
        normal_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        normal_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        normal_format.line_spacing = 1.15
        normal_format.space_after = Pt(8)

        heading_1 = doc.styles["Heading 1"]
        StyleEngine._configure_text_style(
            heading_1,
            font_name="Calibri",
            font_size=16,
            color=theme_color,
            bold=True,
        )
        heading_1.paragraph_format.space_before = Pt(18)
        heading_1.paragraph_format.space_after = Pt(8)

        heading_2 = doc.styles["Heading 2"]
        StyleEngine._configure_text_style(
            heading_2,
            font_name="Calibri",
            font_size=13,
            color=(0, 0, 0),
            bold=True,
        )
        heading_2.paragraph_format.space_before = Pt(14)
        heading_2.paragraph_format.space_after = Pt(6)

        heading_3 = doc.styles["Heading 3"]
        StyleEngine._configure_text_style(
            heading_3,
            font_name="Calibri",
            font_size=12,
            color=(54, 54, 54),
            bold=True,
        )
        heading_3.paragraph_format.space_before = Pt(10)
        heading_3.paragraph_format.space_after = Pt(4)

        for list_style_name in (
            "List Bullet",
            "List Bullet 2",
            "List Bullet 3",
            "List Number",
            "List Number 2",
            "List Number 3",
        ):
            if list_style_name not in doc.styles:
                continue
            list_style = doc.styles[list_style_name]
            StyleEngine._configure_text_style(
                list_style,
                font_name="Calibri",
                font_size=11,
                color=(33, 37, 41),
            )
            list_style.paragraph_format.space_after = Pt(4)


class ChartEngine:
    @staticmethod
    def _get_plt_color(theme_color):
        return tuple(channel / 255 for channel in theme_color)

    @staticmethod
    def create_bar_chart(data_str, theme_color):
        try:
            parts = data_str.split("|")
            if len(parts) == 3:
                title_str, ylabel_str, raw_data = [part.strip() for part in parts]
            else:
                title_str, ylabel_str, raw_data = (
                    "Sentimen Makro",
                    "Persentase",
                    data_str,
                )

            labels, values = [], []
            for pair in raw_data.split(";"):
                if "," not in pair:
                    continue
                label, value = pair.split(",", maxsplit=1)
                cleaned_value = re.sub(r"[^\d.]", "", value)
                if not cleaned_value:
                    continue
                labels.append(label.strip())
                values.append(float(cleaned_value))

            if not labels:
                return None

            fig, axis = plt.subplots(figsize=(7, 4.5))
            axis.bar(
                labels,
                values,
                color=ChartEngine._get_plt_color(theme_color),
                alpha=0.9,
                width=0.5,
            )
            axis.set_title(title_str, fontsize=12, fontweight="bold", pad=20)
            axis.set_ylabel(ylabel_str, fontsize=10)
            axis.spines["top"].set_visible(False)
            axis.spines["right"].set_visible(False)

            image_stream = io.BytesIO()
            plt.savefig(image_stream, format="png", bbox_inches="tight", dpi=150)
            plt.close(fig)
            image_stream.seek(0)
            return image_stream
        except Exception as exc:
            logger.warning("Gagal membuat bar chart: %s", exc)
            return None

    @staticmethod
    def create_flowchart(data_str, theme_color):
        try:
            steps = [
                "\n".join(textwrap.wrap(step.strip(), width=18))
                for step in data_str.split("->")
                if step.strip()
            ]
            if len(steps) < 2:
                return None

            fig, axis = plt.subplots(figsize=(8, 3))
            axis.axis("off")
            x_positions = [index * 2.5 for index in range(len(steps))]

            for index in range(len(steps) - 1):
                axis.annotate(
                    "",
                    xy=(x_positions[index + 1] - 1.0, 0.5),
                    xytext=(x_positions[index] + 1.0, 0.5),
                    arrowprops={"arrowstyle": "-|>", "lw": 1.5},
                )

            for index, step in enumerate(steps):
                box = patches.FancyBboxPatch(
                    (x_positions[index] - 1.0, 0.1),
                    2.0,
                    0.8,
                    boxstyle="round,pad=0.1",
                    fc=ChartEngine._get_plt_color(theme_color),
                    alpha=0.9,
                )
                axis.add_patch(box)
                axis.text(
                    x_positions[index],
                    0.5,
                    step,
                    ha="center",
                    va="center",
                    size=9,
                    color="white",
                    fontweight="bold",
                )

            axis.set_xlim(-1.2, (len(steps) - 1) * 2.5 + 1.2)
            axis.set_ylim(0, 1)

            image_stream = io.BytesIO()
            plt.savefig(
                image_stream,
                format="png",
                bbox_inches="tight",
                dpi=200,
                transparent=True,
            )
            plt.close(fig)
            image_stream.seek(0)
            return image_stream
        except Exception as exc:
            logger.warning("Gagal membuat flowchart: %s", exc)
            return None


class DocumentBuilder:
    LIST_STYLES = {
        "ul": ["List Bullet", "List Bullet 2", "List Bullet 3"],
        "ol": ["List Number", "List Number 2", "List Number 3"],
    }

    @staticmethod
    def _resolve_list_style(doc, list_tag, level):
        style_candidates = DocumentBuilder.LIST_STYLES.get(list_tag, ["List Bullet"])
        style_name = style_candidates[min(level, len(style_candidates) - 1)]
        fallback_style = style_candidates[0]
        return style_name if style_name in doc.styles else fallback_style

    @staticmethod
    def _append_inline_runs(paragraph, node, bold=False, italic=False):
        if isinstance(node, NavigableString):
            text = str(node)
            if not text:
                return
            run = paragraph.add_run(text)
            run.bold = bold
            run.italic = italic
            return

        if not isinstance(node, Tag):
            return

        if node.name == "br":
            paragraph.add_run("\n")
            return

        next_bold = bold or node.name in {"strong", "b"}
        next_italic = italic or node.name in {"em", "i"}

        for child in node.children:
            DocumentBuilder._append_inline_runs(
                paragraph,
                child,
                bold=next_bold,
                italic=next_italic,
            )

    @staticmethod
    def _render_list(doc, list_node, level=0):
        style_name = DocumentBuilder._resolve_list_style(doc, list_node.name, level)

        for list_item in list_node.find_all("li", recursive=False):
            paragraph = doc.add_paragraph(style=style_name)
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

            for child in list_item.contents:
                if isinstance(child, Tag) and child.name in {"ul", "ol"}:
                    continue
                DocumentBuilder._append_inline_runs(paragraph, child)

            nested_lists = [
                child
                for child in list_item.contents
                if isinstance(child, Tag) and child.name in {"ul", "ol"}
            ]
            for nested in nested_lists:
                DocumentBuilder._render_list(doc, nested, level=level + 1)

    @staticmethod
    def _render_table(doc, table_node):
        rows = table_node.find_all("tr")
        if not rows:
            return

        column_count = max(
            len(row.find_all(["th", "td"], recursive=False))
            for row in rows
        )
        table = doc.add_table(rows=0, cols=column_count)
        table.style = "Table Grid"

        for row_index, row_node in enumerate(rows):
            cells = row_node.find_all(["th", "td"], recursive=False)
            table_row = table.add_row().cells

            for col_index in range(column_count):
                text = ""
                if col_index < len(cells):
                    text = cells[col_index].get_text(" ", strip=True)
                table_row[col_index].text = text

                if row_index == 0 and col_index < len(cells) and cells[col_index].name == "th":
                    for run in table_row[col_index].paragraphs[0].runs:
                        run.bold = True

    @staticmethod
    def parse_html_to_docx(doc, html_content):
        soup = BeautifulSoup(html_content, "html.parser")
        for element in soup.contents:
            if not isinstance(element, Tag):
                continue

            if element.name in {"h1", "h2", "h3"}:
                level = int(element.name[1])
                doc.add_heading(element.get_text(" ", strip=True), level=level)
                continue

            if element.name == "p":
                paragraph = doc.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                for child in element.children:
                    DocumentBuilder._append_inline_runs(paragraph, child)
                continue

            if element.name in {"ul", "ol"}:
                DocumentBuilder._render_list(doc, element, level=0)
                continue

            if element.name == "table":
                DocumentBuilder._render_table(doc, element)

    @staticmethod
    def process_content(doc, raw_text, theme_color=DEFAULT_COLOR):
        clean_lines = []
        for line in raw_text.split("\n"):
            stripped_line = line.strip()

            if stripped_line.startswith("[[CHART:") and stripped_line.endswith("]]"):
                chart_data = stripped_line.replace("[[CHART:", "").replace("]]", "").strip()
                image = ChartEngine.create_bar_chart(chart_data, theme_color)
                if image:
                    doc.add_paragraph().add_run().add_picture(image, width=Inches(5.5))
                continue

            if stripped_line.startswith("[[FLOW:") and stripped_line.endswith("]]"):
                flow_data = stripped_line.replace("[[FLOW:", "").replace("]]", "").strip()
                image = ChartEngine.create_flowchart(flow_data, theme_color)
                if image:
                    doc.add_paragraph().add_run().add_picture(image, width=Inches(6.5))
                continue

            clean_lines.append(line)

        html = markdown.markdown("\n".join(clean_lines), extensions=["tables"])
        DocumentBuilder.parse_html_to_docx(doc, html)

    @staticmethod
    def add_table_of_contents(doc):
        doc.add_heading("DAFTAR ISI", level=1)
        toc_paragraph = doc.add_paragraph()
        append_field(toc_paragraph, 'TOC \\o "1-3" \\h \\z \\u')

        note = doc.add_paragraph(
            "Perbarui field di Microsoft Word agar daftar isi otomatis terisi."
        )
        note.runs[0].italic = True
        note.runs[0].font.size = Pt(9)
        note.alignment = WD_ALIGN_PARAGRAPH.LEFT
        doc.add_page_break()

    @staticmethod
    def create_cover(doc, timeframe, theme_color=DEFAULT_COLOR):
        StyleEngine.apply_document_styles(doc, theme_color)

        for _ in range(5):
            doc.add_paragraph()

        confidentiality = doc.add_paragraph("S T R I C T L Y   C O N F I D E N T I A L")
        confidentiality.alignment = WD_ALIGN_PARAGRAPH.CENTER
        confidentiality.runs[0].font.size = Pt(10)
        confidentiality.runs[0].font.color.rgb = RGBColor(128, 128, 128)
        confidentiality.runs[0].font.bold = True

        doc.add_paragraph()

        report_title = doc.add_paragraph("HOLISTIC CUSTOMER EXPERIENCE REPORT")
        report_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        report_title.runs[0].font.name = "Calibri"
        report_title.runs[0].font.size = Pt(20)

        company_name = doc.add_paragraph("INIXINDO JOGJA")
        company_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
        company_name.runs[0].font.name = "Calibri"
        company_name.runs[0].font.bold = True
        company_name.runs[0].font.size = Pt(32)
        company_name.runs[0].font.color.rgb = RGBColor(*theme_color)

        doc.add_paragraph()

        period_text = doc.add_paragraph(f"Periode Evaluasi Laporan: {timeframe}")
        period_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
        period_text.runs[0].font.size = Pt(13)

        generated_on = datetime.now().strftime("%d %B %Y")
        generated_text = doc.add_paragraph(f"Tanggal Generasi: {generated_on}")
        generated_text.alignment = WD_ALIGN_PARAGRAPH.CENTER
        generated_text.runs[0].font.size = Pt(11)
        generated_text.runs[0].font.color.rgb = RGBColor(128, 128, 128)

        for _ in range(8):
            doc.add_paragraph()

        prepared_for = doc.add_paragraph(
            f"Prepared for Executive Board by:\n{WRITER_FIRM_NAME}"
        )
        prepared_for.alignment = WD_ALIGN_PARAGRAPH.CENTER
        prepared_for.runs[0].font.bold = True

        doc.add_page_break()
        DocumentBuilder.add_table_of_contents(doc)


class ReportGenerator:
    def __init__(self, kb_instance):
        self.ollama = Client(host=OLLAMA_HOST)
        self.kb = kb_instance
        self.research_pool = concurrent.futures.ThreadPoolExecutor(max_workers=4)

    def _get_visual_directive(self, chapter):
        visual_type = chapter.get("visual")
        if visual_type == "bar_chart":
            return (
                "Mandatory Visual: [[CHART: Sentimen Lintas Demografi | Persentase | "
                "Positif,60; Netral,20; Negatif,20]]"
            )
        if visual_type == "flowchart":
            return (
                "Action Plan Visual: [[FLOW: Tinjau Keluhan Mayoritas -> "
                "Sinkronisasi Lintas Divisi -> Implementasi Solusi]]."
            )
        return "Do not force visuals."

    def _fetch_chapter_context(
        self,
        chapter,
        timeframe,
        notes,
        macro_trends,
        chapter_trends,
    ):
        try:
            rag_data = self.kb.query(
                timeframe,
                f"{chapter['focus_keywords']} {notes}".strip(),
            )
            persona = PERSONAS.get("default", "Chief CX Officer")
            section_markdown = "\n".join(
                f"### {section}" for section in chapter["sections"]
            )

            prompt = CX_ANALYSIS_SYSTEM_PROMPT.format(
                persona=persona,
                timeframe=timeframe,
                industry_trends=macro_trends,
                chapter_osint=chapter_trends,
                rag_data=rag_data,
                visual_prompt=self._get_visual_directive(chapter),
                chapter_title=chapter["title"],
                sub_chapters=section_markdown,
            )
            return {"prompt": prompt, "success": True}
        except Exception as exc:
            return {"prompt": "", "success": False, "error": str(exc)}

    def run(self, timeframe, notes=""):
        logger.info("Starting Holistic CX generation for timeframe: %s", timeframe)

        macro_future = self.research_pool.submit(
            Researcher.get_macro_trends,
            timeframe,
            notes,
        )
        topic_futures = {
            chapter["id"]: self.research_pool.submit(
                Researcher.get_topic_trends,
                timeframe,
                chapter["focus_keywords"],
                notes,
            )
            for chapter in CX_SENTIMENT_STRUCTURE
        }

        try:
            macro_trends = macro_future.result(timeout=25)
        except Exception:
            macro_trends = "Tidak ada tren eksternal yang berhasil dimuat."

        chapter_trends = {}
        for chapter_id, future in topic_futures.items():
            try:
                chapter_trends[chapter_id] = future.result(timeout=15)
            except Exception:
                chapter_trends[chapter_id] = "Tidak ada sinyal OSINT fokus bab."

        document = Document()
        DocumentBuilder.create_cover(document, timeframe, DEFAULT_COLOR)

        for index, chapter in enumerate(CX_SENTIMENT_STRUCTURE):
            context = self._fetch_chapter_context(
                chapter,
                timeframe,
                notes,
                macro_trends,
                chapter_trends.get(chapter["id"], ""),
            )
            if not context["success"]:
                logger.error(
                    "Gagal membangun konteks untuk bab %s: %s",
                    chapter["title"],
                    context.get("error", "unknown error"),
                )
                continue

            try:
                response = self.ollama.chat(
                    model=LLM_MODEL,
                    messages=[
                        {"role": "system", "content": context["prompt"]},
                        {
                            "role": "user",
                            "content": (
                                f"Write content for {chapter['title']}. "
                                "Remember: Use '###' for every sub-chapter header and compare "
                                "demographics holistically."
                            ),
                        },
                    ],
                    options={"num_ctx": 4096},
                )
                document.add_heading(chapter["title"], level=1)
                DocumentBuilder.process_content(
                    document,
                    response["message"]["content"],
                    DEFAULT_COLOR,
                )
                if index < len(CX_SENTIMENT_STRUCTURE) - 1:
                    document.add_page_break()
            except Exception as exc:
                logger.error("Error saat memproses %s: %s", chapter["title"], exc)

        filename = f"Inixindo_Holistic_CX_Report_{timeframe}".replace(" ", "_")
        return document, filename
