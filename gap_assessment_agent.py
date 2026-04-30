import io
import json
import os
import streamlit as st
from openai import OpenAI
from docx import Document
import PyPDF2
import pandas as pd

from pptx import Presentation
from pptx.util import Inches, Pt


# --------------------
# Page Setup
# --------------------
st.set_page_config(page_title="AI Gap Assessment Builder", layout="wide")

st.title("Analytics Modernization Assessment Copilot")
st.caption("Upload discovery inputs and generate a client-ready Word gap assessment.")


# --------------------
# OpenAI Setup
# --------------------
api_key = st.secrets.get("OPENAI_API_KEY", None) or os.getenv("OPENAI_API_KEY")

if not api_key:
    st.error("OPENAI_API_KEY is missing. Add it in Streamlit Cloud Secrets.")
    st.stop()

client = OpenAI(api_key=api_key)


# --------------------
# UI Inputs
# --------------------
client_name = st.text_input("Client Name")
industry = st.text_input("Industry")

assessment_type = st.selectbox(
    "Assessment Type",
    [
        "Analytics Gap Assessment",
        "SAP / S/4HANA Reporting Assessment",
        "AI Opportunity Assessment",
        "Data Strategy Assessment"
    ]
)

uploaded_files = st.file_uploader(
    "Upload Discovery Notes / Supporting Files",
    type=["txt", "csv", "pdf", "xls", "xlsx"],
    accept_multiple_files=True
)

notes = st.text_area("Paste Additional Notes", height=250)


# --------------------
# File Reader
# --------------------
def read_uploaded_files(files):
    content = ""

    if not files:
        return content

    for file in files:
        content += f"\n\n--- FILE: {file.name} ---\n"

        file_name = file.name.lower()

        try:
            if file_name.endswith(".txt") or file.type == "text/plain":
                content += file.read().decode("utf-8", errors="ignore")

            elif file_name.endswith(".csv") or file.type == "text/csv":
                df = pd.read_csv(file)
                content += df.head(25).to_string(index=False)

            elif file_name.endswith(".xlsx"):
                excel_file = pd.ExcelFile(file, engine="openpyxl")
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    content += f"\n\n--- SHEET: {sheet_name} ---\n"
                    content += df.head(20).to_string(index=False)

            elif file_name.endswith(".xls"):
                excel_file = pd.ExcelFile(file, engine="xlrd")
                for sheet_name in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    content += f"\n\n--- SHEET: {sheet_name} ---\n"
                    content += df.head(20).to_string(index=False)

            elif file_name.endswith(".pdf") or file.type == "application/pdf":
                reader = PyPDF2.PdfReader(file)
                for i, page in enumerate(reader.pages, start=1):
                    content += f"\n\n--- PAGE {i} ---\n"
                    content += page.extract_text() or ""

            else:
                content += f"\nUnsupported file type: {file.type}"

        except Exception as e:
            content += f"\nError reading file: {str(e)}"

    return content


# --------------------
# OpenAI Retry Helper
# --------------------
from openai import RateLimitError, APIError, APITimeoutError
import time
import json

def call_openai_with_retry(messages, model="gpt-4o-mini"):
    for attempt in range(3):
        try:
            return client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=0.2,
                max_tokens=2500
            )

        except RateLimitError:
            if attempt < 2:
                time.sleep(2 ** attempt)
            else:
                st.error("OpenAI rate limit reached. Check billing/quota or reduce upload size.")
                return None

        except (APIError, APITimeoutError) as e:
            st.error(f"OpenAI API error: {str(e)}")
            return None


# --------------------
# Generate Assessment JSON
# --------------------
def generate_assessment_json(client_name, industry, assessment_type, notes, file_content):

    notes = notes[:4000]
    file_content = file_content[:12000]
    
    prompt = f"""
You are a senior enterprise consulting partner creating an analytics gap assessment.

Client Name: {client_name}
Industry: {industry}
Assessment Type: {assessment_type}

Discovery Notes:
{notes}

Supporting File Content:
{file_content}

Return ONLY valid JSON.

Create JSON with these exact keys:

engagement_overview_text
engagement_scope_summary
executive_summary_text
analytics_environment_snapshot
analytics_complexity_text
analytics_complexity_snapshot
gap_heatmap_intro
gap_severity_heatmap
gap_observations_text
current_landscape_text
current_architecture_summary
reporting_inventory_text
reporting_landscape_summary
s4_reporting_impact_text
s4_impact_summary
key_gaps_text
gap_analysis_summary
opportunity_areas_text
improvement_opportunity_summary
business_value_text
potential_impact_summary
recommended_next_steps_text
recommended_focus_areas
appendix_reporting_inventory
appendix_s4_impact_analysis
appendix_reporting_overlap_analysis
appendix_data_source_mapping
appendix_critical_reports
critical_report_summary
analytics_ownership_overview
analytics_responsibility_model
stakeholder_interview_summary
responsibility_gaps
key_observations_text

Rules:
- Return ONLY valid JSON. No markdown.
- Every table field must be an array of objects.
- Do not return nested objects inside table fields.
- Do not return Python-style lists as strings.
- Do not include S/4HANA content unless SAP, ECC, or S/4HANA is mentioned in the notes.
- For non-SAP clients, set S/4HANA sections to "Not applicable based on current discovery inputs."
- Avoid generic consulting language.
- Tie every gap and recommendation to the client facts.
- Use business-friendly language for executives.
"""

    messages = [
        {"role": "system", "content": "Return only valid JSON. No markdown."},
        {"role": "user", "content": prompt}
    ]

    response = call_openai_with_retry(messages)

    if response is None:
        return {}

    raw = response.choices[0].message.content.strip()

    if raw.startswith("```"):
        raw = raw.replace("```json", "").replace("```", "").strip()

    return json.loads(raw)

# --------------------
# Word Helpers
# --------------------
def add_heading(doc, text, level=1):
    doc.add_heading(text, level=level)


def add_paragraph(doc, text):
    if not text:
        doc.add_paragraph("To be validated.")
        return

    if isinstance(text, dict):
        text = json.dumps(text, indent=2)
    elif isinstance(text, list):
        text = "\n".join([str(item) for item in text])
    else:
        text = str(text)

    doc.add_paragraph(text)


def add_table_from_records(doc, records):
    if not records:
        doc.add_paragraph("To be validated.")
        return

    if isinstance(records, str):
        doc.add_paragraph(records)
        return

    if isinstance(records, dict):
        records = [records]

    if not isinstance(records, list) or len(records) == 0:
        doc.add_paragraph("To be validated.")
        return

    if isinstance(records[0], str):
        for item in records:
            doc.add_paragraph(str(item), style="List Bullet")
        return

    if not isinstance(records[0], dict):
        doc.add_paragraph(str(records))
        return

    headers = list(records[0].keys())

    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"

    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = str(h)

    for record in records:
        row = table.add_row().cells

        for i, h in enumerate(headers):
            value = record.get(h, "")

            if isinstance(value, dict):
                value = json.dumps(value, indent=2)
            elif isinstance(value, list):
                value = ", ".join([str(x) for x in value])
            else:
                value = str(value)

            row[i].text = value


# --------------------
# Build Word Document
# --------------------
def build_docx(data, client_name):
    doc = Document()

    doc.add_heading(f"{client_name or 'Client'} Analytics Gap Assessment", 0)

    add_heading(doc, "1. Engagement Overview", 1)
    add_paragraph(doc, data.get("engagement_overview_text", ""))
    add_heading(doc, "Engagement Scope Summary", 2)
    add_table_from_records(doc, data.get("engagement_scope_summary", []))

    add_heading(doc, "2. Executive Summary", 1)
    add_paragraph(doc, data.get("executive_summary_text", ""))
    add_heading(doc, "Analytics Environment Snapshot", 2)
    add_table_from_records(doc, data.get("analytics_environment_snapshot", []))

    add_heading(doc, "3. Analytics Complexity Snapshot", 1)
    add_paragraph(doc, data.get("analytics_complexity_text", ""))
    add_table_from_records(doc, data.get("analytics_complexity_snapshot", []))

    add_heading(doc, "4. Analytics Gap Severity Heatmap", 1)
    add_paragraph(doc, data.get("gap_heatmap_intro", ""))
    add_table_from_records(doc, data.get("gap_severity_heatmap", []))
    add_heading(doc, "Observations", 2)
    add_paragraph(doc, data.get("gap_observations_text", ""))

    add_heading(doc, "5. Current Analytics Landscape", 1)
    add_paragraph(doc, data.get("current_landscape_text", ""))
    add_heading(doc, "Current Analytics Architecture Summary", 2)
    add_table_from_records(doc, data.get("current_architecture_summary", []))

    add_heading(doc, "6. Reporting Inventory Summary", 1)
    add_paragraph(doc, data.get("reporting_inventory_text", ""))
    add_heading(doc, "Reporting Landscape Summary", 2)
    add_table_from_records(doc, data.get("reporting_landscape_summary", []))

    add_heading(doc, "7. S/4HANA Reporting Impact Assessment", 1)
    add_paragraph(doc, data.get("s4_reporting_impact_text", ""))
    add_heading(doc, "S/4HANA Reporting Impact Summary", 2)
    add_table_from_records(doc, data.get("s4_impact_summary", []))

    add_heading(doc, "8. Key Analytics Gaps Identified", 1)
    add_paragraph(doc, data.get("key_gaps_text", ""))
    add_heading(doc, "Gap Analysis Summary", 2)
    add_table_from_records(doc, data.get("gap_analysis_summary", []))

    add_heading(doc, "9. Opportunity Areas", 1)
    add_paragraph(doc, data.get("opportunity_areas_text", ""))
    add_heading(doc, "Improvement Opportunity Summary", 2)
    add_table_from_records(doc, data.get("improvement_opportunity_summary", []))

    add_heading(doc, "10. Business Value of Addressing the Gaps", 1)
    add_paragraph(doc, data.get("business_value_text", ""))
    add_heading(doc, "Potential Impact", 2)
    add_table_from_records(doc, data.get("potential_impact_summary", []))

    add_heading(doc, "11. Recommended Next Steps", 1)
    add_paragraph(doc, data.get("recommended_next_steps_text", ""))
    add_heading(doc, "Recommended Focus Areas", 2)
    add_table_from_records(doc, data.get("recommended_focus_areas", []))

    add_heading(doc, "12. Appendix A — Reporting Inventory", 1)
    add_table_from_records(doc, data.get("appendix_reporting_inventory", []))

    add_heading(doc, "13. Appendix B — S/4 Reporting Impact Analysis", 1)
    add_table_from_records(doc, data.get("appendix_s4_impact_analysis", []))

    add_heading(doc, "14. Appendix C — Reporting Overlap Analysis", 1)
    add_table_from_records(doc, data.get("appendix_reporting_overlap_analysis", []))

    add_heading(doc, "15. Appendix D — Data Source Mapping Table", 1)
    add_table_from_records(doc, data.get("appendix_data_source_mapping", []))

    add_heading(doc, "16. Appendix E — Critical Reports", 1)
    add_table_from_records(doc, data.get("appendix_critical_reports", []))

    add_heading(doc, "Critical Report Summary", 2)
    add_table_from_records(doc, data.get("critical_report_summary", []))

    add_heading(doc, "17. Appendix F — Analytics Stakeholder Map", 1)

    add_heading(doc, "Analytics Ownership Overview", 2)
    add_table_from_records(doc, data.get("analytics_ownership_overview", []))

    add_heading(doc, "Analytics Responsibility Model", 2)
    add_table_from_records(doc, data.get("analytics_responsibility_model", []))

    add_heading(doc, "Stakeholder Interview Summary", 2)
    add_table_from_records(doc, data.get("stakeholder_interview_summary", []))

    add_heading(doc, "Key Analytics Responsibility Gaps", 2)
    add_table_from_records(doc, data.get("responsibility_gaps", []))

    add_heading(doc, "Key Observations", 2)
    add_paragraph(doc, data.get("key_observations_text", ""))

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    return output

def build_ppt(data, client_name):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Slide 1 Title
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = f"{client_name} Analytics Gap Assessment"

    tx = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(12), Inches(4))
    tf = tx.text_frame
    p = tf.add_paragraph()
    p.text = data.get("executive_summary_text", "Executive summary unavailable.")
    p.font.size = Pt(18)

    # Slide 2 Key Gaps
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Key Analytics Gaps"

    tx = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(12), Inches(5))
    tf = tx.text_frame

    gaps = data.get("gap_analysis_summary", [])

    for gap in gaps[:6]:
        p = tf.add_paragraph()
        p.text = f"• {gap.get('Gap','Gap')} – {gap.get('Business Impact','')}"
        p.font.size = Pt(16)

    # Slide 3 Recommendations
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Recommended Next Steps"

    tx = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(12), Inches(5))
    tf = tx.text_frame

    focus = data.get("recommended_focus_areas", [])

    for item in focus[:6]:
        p = tf.add_paragraph()
        p.text = f"• {item.get('Focus Area','')} – {item.get('Recommended Next Step','')}"
        p.font.size = Pt(16)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    return output

def build_exec_email(data, client_name):
    summary = data.get("executive_summary_text", "")
    
    email = f"""
Subject: {client_name} Analytics Gap Assessment – Executive Summary

Team,

We completed the initial analytics gap assessment for {client_name}.

Key observations:
{summary}

Top priorities identified:
1. Centralize reporting and KPI visibility
2. Improve data integration across systems
3. Enable forecasting and operational analytics
4. Build scalable analytics foundation for growth

Recommended next step:
Conduct a focused strategy workshop and roadmap session to prioritize quick wins and transformation initiatives.

Regards,
Consulting Team
"""
    return email


# --------------------
# Generate Button
# --------------------
if st.button("Generate Assessment Outputs"):
    if not client_name:
        st.warning("Enter a client name first.")
    else:
        file_content = read_uploaded_files(uploaded_files)

        with st.spinner("Generating assessment content..."):
            data = generate_assessment_json(
                client_name,
                industry,
                assessment_type,
                notes,
                file_content
            )

        safe_client_name = (
            client_name.replace(" ", "_")
            .replace("'", "")
            .replace("/", "_")
        )

        st.success("Assessment generated.")

        # --------------------
        # Word Document
        # --------------------
        with st.spinner("Creating Word document..."):
            docx_file = build_docx(data, client_name)

        st.download_button(
            label="Download Word Assessment",
            data=docx_file.getvalue(),
            file_name=f"{safe_client_name}_Gap_Assessment.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

        # --------------------
        # PowerPoint Deck
        # --------------------
        with st.spinner("Creating PowerPoint deck..."):
            ppt_file = build_ppt(data, client_name)

        st.download_button(
            label="Download PowerPoint Deck",
            data=ppt_file.getvalue(),
            file_name=f"{safe_client_name}_Gap_Assessment_Deck.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

        # --------------------
        # Executive Summary Email
        # --------------------
        with st.spinner("Creating Executive Summary Email..."):
            email_text = build_exec_email(data, client_name)

        st.download_button(
            label="Download Executive Summary Email",
            data=email_text,
            file_name=f"{safe_client_name}_Executive_Summary_Email.txt",
            mime="text/plain"
        )
