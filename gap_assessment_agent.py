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
from tavily import TavilyClient


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

#----------------------
# Tavili Setup
#______________________



tavily_api_key = st.secrets.get("TAVILY_API_KEY", None) or os.getenv("TAVILY_API_KEY")

tavily_client = None
if tavily_api_key:
    tavily_client = TavilyClient(api_key=tavily_api_key)


# --------------------
# Initialize Session State
# --------------------
defaults = {
    "assessment_data": None,
    "word_doc": None,
    "ppt_file": None,
    "email_text": None,
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

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

safe_client_name = client_name.strip().replace(" ", "_") if client_name else "Client"


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

response_format={"type": "json_object"}

def call_openai_with_retry(messages, model="gpt-4o-mini"):
    for attempt in range(3):
        try:
            return client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=0.2,
                max_tokens=4000,
                response_format={"type": "json_object"}
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
# Research Company Information
# --------------------

def research_company(company_name, industry):
    if not tavily_client or not company_name:
        return ""

    query = f"{company_name} company overview industry products revenue locations acquisitions strategy {industry}"

    try:
        results = tavily_client.search(
            query=query,
            search_depth="basic",
            max_results=5
        )

        research_text = ""

        for item in results.get("results", []):
            title = item.get("title", "")
            url = item.get("url", "")
            content = item.get("content", "")

            research_text += f"\nTitle: {title}\nURL: {url}\nSummary: {content}\n"

        return research_text[:6000]

    except Exception as e:
        return f"Company research unavailable: {str(e)}"

# --------------------
# Generate Assessment JSON
# --------------------
def generate_assessment_json(client_name, industry, assessment_type, notes, file_content, company_research):

    notes = notes[:4000]
    file_content = file_content[:12000]

    prompt = f"""
You are a senior consulting partner from a top-tier advisory firm delivering a paid executive assessment for a client.

Your writing style must feel premium, strategic, commercial, and boardroom-ready.

Never sound generic, robotic, repetitive, or AI-generated.

The final output must feel like a deliverable a client would pay $50,000+ for.

CLIENT INFORMATION

Client Name: {client_name}
Industry: {industry}
Assessment Type: {assessment_type}

PUBLIC COMPANY RESEARCH

{company_research}

DISCOVERY NOTES

{notes}

SUPPORTING FILE CONTENT

{file_content}


OBJECTIVE

Create a premium executive analytics assessment in JSON format.

The document must first explain the company, what it has accomplished, where complexity has increased, and why leadership requested this assessment now.

Then identify reporting, analytics, governance, technology, and decision-support gaps.

Then recommend practical next steps tied to business value.

WRITING REQUIREMENTS

Write like an experienced consulting executive.

Use specific business language such as:
growth, margin pressure, operating visibility, scalability, decision-making speed, reporting trust, working capital, service levels, operational efficiency, transformation readiness.

Tie all observations directly to likely client realities.

For manufacturing clients, naturally reference:
plants, supply chain, production, inventory, OTIF, forecasting, downtime, scrap, yield, procurement, distribution, acquisitions.

All tables must be rendered as clean tabular structures with column headers and rows.

Do NOT return Python dictionaries or JSON-style inline objects.

Each table must be formatted as an array of flat row objects that can be directly rendered into a professional table.

Each row must contain:
- Business context
- Impact
- Priority
- Action (where applicable)

Do not use generic columns like "Category" or "Gap".

Each row must read like a consulting insight, not a label.

MANDATORY POINT OF VIEW

You are not summarizing findings.

You are diagnosing the business.

You must take a clear position on:

- What is actually broken
- Where it shows up operationally
- Why it matters financially
- What leadership should prioritize first

Avoid neutral language.

Write with conviction as if presenting to a CFO and COO.

Avoid generic labels like "gaps" or "opportunities" as single columns.

Do NOT use generic filler statements such as:
"significant opportunities exist"
"the company faces challenges"
"there are several gaps"

Be specific and commercial.

Do not create nested dictionaries inside table cells.

Every table must be a list of flat row objects.

Bad:
[
  {{"stakeholders": [{{"name": "CFO", "role": "Finance"}}]}}
]

Good:
[
  {{"Stakeholder": "CFO", "Role": "Finance", "Current Pain Point": "Delayed margin visibility", "Business Risk": "Slow pricing decisions", "Requested Capability": "Weekly profitability dashboard", "Priority": "High"}}
]

Use clean column names with spaces and title case.
Never use one-column tables.
Never return fields like "gaps", "opportunities", "ownership", "interviews", or "reporting_inventory" as a single nested column.

MANDATORY COMPLETENESS RULE

Every required key must contain meaningful content.

- No empty arrays
- No empty strings
- No "To be validated"
- No placeholder text

If information is not available, infer realistic consulting-level content.

Do not skip sections under any circumstance.

CRITICAL REQUIREMENT – ROADMAP

The implementation_roadmap section is mandatory.

You MUST return:

Phase 1 (0–6 weeks)
Phase 2 (6–12 weeks)
Phase 3 (12+ weeks)

Each phase must include:
- Key Actions
- Business Outcomes
- Dependencies

If this section is missing, the response is invalid.

BUSINESS TRANSLATION REQUIREMENT

Every issue identified must include:

1. Where it shows up in the business (function/process)
2. What is happening today
3. The consequence (delay, inefficiency, cost, risk)
4. Why it matters (margin, cash flow, service levels, growth)

Do not describe issues without tying them to real business impact.

ENGAGEMENT OVERVIEW MUST INCLUDE:

1. What the company does
2. What it has accomplished / growth journey
3. Why complexity increased
4. Why now is the right time for assessment
5. Why analytics matters to leadership now

EXECUTIVE SUMMARY MUST INCLUDE:

1. Current state reality
2. Main risks
3. Business consequences
4. Biggest value opportunities
5. Immediate recommended actions

TABLE STRUCTURE REQUIREMENT

All tables must follow this structure:

Business Area
Current State
Where It Breaks
Business Impact
Why It Matters
Recommended Action
Priority

TOP PRIORITIES REQUIREMENT

Identify the top 5 actions leadership should take.

Each must include:

- What to do
- Why it matters
- Business impact
- Time horizon (Immediate / Near-Term / Mid-Term)

These should feel like executive decisions, not suggestions.

ROADMAP REQUIREMENT

Create a phased roadmap:

Phase 1 (0–6 weeks): Quick wins and stabilization
Phase 2 (6–12 weeks): Foundation build
Phase 3 (12+ weeks): Scale and advanced capabilities

Each phase must include:
- Key actions
- Business outcomes
- Dependencies

VALUE QUANTIFICATION REQUIREMENT

Where possible, estimate impact using directional ranges:

- Time reduction (e.g., 20–40%)
- Cost savings
- Margin improvement
- Working capital impact

Do not invent unrealistic numbers.

Use reasonable, experience-based estimates tied to the issue.

RETURN ONLY VALID JSON

STRICT TABLE FORMAT

All table outputs MUST be arrays of objects using consistent columns.

Do not switch formats between sections.

Do not return narrative where a table is expected.

Inconsistent structure is not allowed.

Required keys:

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
top_priorities
implementation_roadmap

RULES

- Return valid JSON only
- No markdown
- No code fences
- No empty sections
- If data is missing, infer realistic executive-quality content
- Every table field must be an array of objects
- Narrative sections must be premium quality and client-ready
- Recommendations must be practical, phased, and tied to ROI
- Avoid repeating the same wording
- Sound like a paid consulting advisor
"""

    messages = [
        {
            "role": "system",
            "content": "You must return a single valid JSON object only. No markdown, no commentary, no code fences."
        },
        {
            "role": "user",
            "content": prompt
        }
    ]

    response = call_openai_with_retry(messages)

    if response is None:
        return {}

    raw = response.choices[0].message.content.strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        st.error("The AI response was not valid JSON. Showing raw response for debugging:")
        st.code(raw[:4000])
        return {}

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
    p = tf.paragraphs[0]
    p.text = data.get("executive_summary_text", "Executive summary unavailable.")
    p.font.size = Pt(18)

    # Slide 2 Key Gaps
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Key Analytics Gaps"

    tx = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(12), Inches(5))
    tf = tx.text_frame

    gaps = data.get("gap_analysis_summary", [])

    if isinstance(gaps, str):
        gaps = [{"Gap": gaps, "Business Impact": ""}]
    elif isinstance(gaps, dict):
        gaps = [gaps]
    elif not isinstance(gaps, list):
        gaps = []

    for gap in gaps[:6]:
        p = tf.add_paragraph()

        if isinstance(gap, dict):
            p.text = f"• {gap.get('Gap', 'Gap')} – {gap.get('Business Impact', '')}"
        else:
            p.text = f"• {str(gap)}"

        p.font.size = Pt(16)

    # Slide 3 Recommendations
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    slide.shapes.title.text = "Recommended Next Steps"

    tx = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(12), Inches(5))
    tf = tx.text_frame

    focus = data.get("recommended_focus_areas", [])

    if isinstance(focus, str):
        focus = [{"Focus Area": focus, "Recommended Next Step": ""}]
    elif isinstance(focus, dict):
        focus = [focus]
    elif not isinstance(focus, list):
        focus = []

    for item in focus[:6]:
        p = tf.add_paragraph()

        if isinstance(item, dict):
            p.text = f"• {item.get('Focus Area', '')} – {item.get('Recommended Next Step', '')}"
        else:
            p.text = f"• {str(item)}"

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
# Company Research Default
# --------------------
company_research = ""


# --------------------
# Output Validation
# --------------------
def validate_output(data):
    required_keys = [
        "executive_summary_text",
        "top_priorities",
        "implementation_roadmap"
    ]

    for key in required_keys:
        if key not in data or not data[key]:
            return False

    return True

# --------------------
# Generate Button
# --------------------
if st.button("Generate Assessment Outputs", key="main_generate_btn"):

    if not client_name:
        st.warning("Enter a client name first.")
    else:
        file_content = read_uploaded_files(uploaded_files)

        with st.spinner("Generating assessment content..."):

        max_retries = 2
        data = None

    for attempt in range(max_retries + 1):
        data = generate_assessment_json(
            client_name,
            industry,
            assessment_type,
            notes,
            file_content,
            company_research
        )

        if validate_output(data):
            break
        else:
            st.warning(f"Regenerating output (attempt {attempt + 1}) due to missing sections...")

    if not validate_output(data):
        st.error("Failed to generate complete assessment after retries.")
        data = {}

        st.session_state.assessment_data = data

        if data:
            with st.spinner("Creating Word document..."):
                st.session_state.word_doc = build_docx(data, client_name)

            with st.spinner("Creating PowerPoint deck..."):
                st.session_state.ppt_file = build_ppt(data, client_name)

            with st.spinner("Creating Executive Summary Email..."):
                st.session_state.email_text = build_exec_email(data, client_name)

            st.success("Assessment outputs generated successfully.")
        else:
            st.error("Assessment generation failed.")

# --------------------
# Download Buttons
# --------------------
if st.session_state.get("word_doc"):
    st.download_button(
        label="Download Word Document",
        data=st.session_state.word_doc.getvalue() if hasattr(st.session_state.word_doc, "getvalue") else st.session_state.word_doc,
        file_name=f"{safe_client_name}_Gap_Assessment_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        key="download_word_doc"
    )

if st.session_state.get("ppt_file"):
    st.download_button(
        label="Download PowerPoint Deck",
        data=st.session_state.ppt_file.getvalue() if hasattr(st.session_state.ppt_file, "getvalue") else st.session_state.ppt_file,
        file_name=f"{safe_client_name}_Gap_Assessment_Deck.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        key="download_ppt_file"
    )

if st.session_state.get("email_text"):
    st.download_button(
        label="Download Executive Summary Email",
        data=st.session_state.email_text,
        file_name=f"{safe_client_name}_Executive_Summary_Email.txt",
        mime="text/plain",
        key="download_exec_email"
    )
