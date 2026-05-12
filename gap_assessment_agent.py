import json
import os
import streamlit as st
from openai import OpenAI
from docx import Document
import PyPDF2
import pandas as pd
import io

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
        "Analytics Modernization Roadmap",
        "AI Opportunity Assessment"
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
# Shared Prompt Builder
# --------------------
def build_base_context(client_name, industry, assessment_type, notes, file_content, company_research):
    return f"""
CLIENT INFORMATION
Client Name: {client_name}
Industry: {industry}
Assessment Type: {assessment_type}

PUBLIC COMPANY RESEARCH
{company_research}

DISCOVERY NOTES
{notes[:4000]}

SUPPORTING FILE CONTENT
{file_content[:12000]}

{COMPANY_RESEARCH_REQUIREMENTS}

{TABLE_SUMMARY_REQUIREMENTS}



WRITING RULES
You are a senior consulting partner creating a paid executive deliverable.

Write with a direct, commercial, boardroom-ready point of view.

Before writing, infer the company's operating model from the discovery notes, uploaded files, research, and industry.

Do not write generic statements that could apply to any company.

Every issue must explain:
- where it shows up in the business
- what is happening operationally
- who is impacted
- what decision is delayed, wrong, or harder to make
- why it matters financially or operationally

No placeholders.
No "To be validated."
No empty strings.
No empty arrays.
No markdown.
No code fences.

Return valid JSON only.

All tables must be arrays of flat row objects.
Do not create nested dictionaries inside table cells.
"""

# --------------------
# Assessment Context
# --------------------
ASSESSMENT_CONTEXT = {
    "Analytics Gap Assessment": """
You are creating a premium executive-level Analytics Gap Assessment.

The customer is currently undergoing or planning an SAP S/4HANA transformation initiative. The purpose is to evaluate the current-state data, analytics, reporting, governance, ownership, and source-system environment to identify gaps, risks, and business disruption areas caused by source and reporting changes.

Focus on:
- current-state reporting risks
- S/4HANA source-system impact
- report breakage risk
- data model and KPI disruption
- historical reporting challenges
- governance and ownership gaps
- business decision-making risk
""",

    "Analytics Modernization Roadmap": """
You are creating a premium executive-level Analytics Modernization Roadmap.

The customer is working toward an enterprise analytics modernization journey. The purpose is to define the future-state analytics vision, platform strategy, governance model, modernization priorities, roadmap, and business value.

Focus on:
- future-state analytics architecture
- platform consolidation
- reporting modernization
- governance maturity
- self-service analytics
- cloud/data platform strategy
- phased execution roadmap
- business value realization
""",

    "AI Opportunity Assessment": """
You are creating a premium executive-level AI Opportunity Assessment.

The customer is looking to understand AI readiness, practical AI opportunities, and where AI can create value across operations, analytics, automation, and decision-making.

Focus on:
- AI readiness
- data maturity
- AI governance
- automation opportunities
- decision-support use cases
- practical GenAI opportunities
- implementation complexity
- phased AI roadmap
- business value realization
"""
}

COMPANY_RESEARCH_REQUIREMENTS = """
Use the public company research, discovery notes, uploaded files, and industry context to make the assessment company-specific.

Explain the customer's operating model, likely business complexity, reporting needs, ERP/data landscape, and transformation drivers.

Avoid generic language. Do not fabricate specific counts, revenue, locations, systems, reports, or years of history unless provided.
"""

TABLE_SUMMARY_REQUIREMENTS = """
After every table, generate a 1-2 paragraph executive narrative explaining what the table means to the customer.

The narrative must explain:
- why leadership should care
- operational implications
- business risk
- reporting or analytics impact
- financial or decision-making impact
- recommended action

Do not simply repeat the table. Interpret the findings.
"""

# --------------------
# Generate One Section
# --------------------
def generate_section_json(
    client_name,
    industry,
    assessment_type,
    notes,
    file_content,
    company_research,
    section_config,
    assessment_context
):
    base_context = build_base_context(
        client_name,
        industry,
        assessment_type,
        notes,
        file_content,
        company_research
    )

    required_keys = "\n".join(section_config["keys"])

    prompt = f"""
{assessment_context}

{COMPANY_RESEARCH_REQUIREMENTS}

{TABLE_SUMMARY_REQUIREMENTS}

{base_context}

SECTION TO GENERATE:
{section_config["section_name"]}

SECTION INSTRUCTIONS:
{section_config["instructions"]}

REQUIRED JSON KEYS:
{required_keys}

Return only these keys in one valid JSON object.

Requirements:
- Every required key must be populated
- No placeholder text
- No generic consulting language
- Use company-specific business context
- Tie findings to operational and reporting impacts
- Explain business risks and executive implications
- Explain S/4HANA impacts where relevant
- Interpret findings like an executive consultant
"""
    
    messages = [
        {
            "role": "system",
            "content": "Return one valid JSON object only. No markdown. No commentary."
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
        st.error(f"Invalid JSON returned for section: {section_config['section_name']}")
        st.code(raw[:4000])
        return {}


# --------------------
# Generate Full Assessment - Multiple Calls
# --------------------
def generate_assessment_json(
    client_name,
    industry,
    assessment_type,
    notes,
    file_content,
    company_research
):
    framework = ASSESSMENT_FRAMEWORKS.get(assessment_type)
    assessment_context = ASSESSMENT_CONTEXT.get(assessment_type, "")

    if not framework:
        st.error(f"No framework found for assessment type: {assessment_type}")
        return {}

    assessment_context = ASSESSMENT_CONTEXT.get(assessment_type, "")

    final_data = {}

    for section in framework["sections"]:
        with st.spinner(f"Generating {section['section_name']}..."):
            section_data = generate_section_json(
                client_name,
                industry,
                assessment_type,
                notes,
                file_content,
                company_research,
                section,
                assessment_context
            )

        final_data.update(section_data)

    return final_data

# --------------------
# Word Helpers
# --------------------
def add_heading(doc, text, level=1):
    doc.add_heading(text, level=level)


def add_paragraph(doc, text):
    if not text:
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
def build_docx(data, client_name, assessment_type):
    doc = Document()

    doc.add_heading(f"{client_name or 'Client'} {assessment_type}", 0)

    # Common executive front-end
    add_heading(doc, "1. Engagement Overview", 1)
    add_paragraph(doc, data.get("engagement_overview_text", ""))

    add_heading(doc, "2. Executive Summary", 1)
    add_paragraph(doc, data.get("executive_summary_text", ""))

    if data.get("top_priorities"):
        add_heading(doc, "Executive Priorities", 2)
        add_table_from_records(doc, data.get("top_priorities", []))

    if data.get("implementation_roadmap"):
        add_heading(doc, "Implementation Roadmap", 2)
        add_table_from_records(doc, data.get("implementation_roadmap", []))


    

    # --------------------
    # Analytics Gap Assessment
    # --------------------
    if assessment_type == "Analytics Gap Assessment":

        add_heading(doc, "3. Analytics Environment Snapshot", 1)
        add_table_from_records(doc, data.get("analytics_environment_snapshot", []))
        add_paragraph(doc, data.get("analytics_environment_summary", ""))

        add_heading(doc, "4. Analytics Complexity Snapshot", 1)
        add_paragraph(doc, data.get("analytics_complexity_text", ""))
        add_table_from_records(doc, data.get("analytics_complexity_snapshot", []))

        add_heading(doc, "5. Gap Severity Heatmap", 1)
        add_table_from_records(doc, data.get("gap_severity_heatmap", []))
        add_paragraph(doc, data.get("gap_observations_text", ""))

        add_heading(doc, "6. Current Analytics Landscape", 1)
        add_paragraph(doc, data.get("current_landscape_text", ""))
        add_table_from_records(doc, data.get("current_architecture_summary", []))

        add_heading(doc, "7. Reporting Inventory Summary", 1)
        add_table_from_records(doc, data.get("reporting_landscape_summary", []))
        add_paragraph(doc, data.get("reporting_inventory_text", ""))

        add_heading(doc, "8. S/4HANA Reporting Impact", 1)
        add_table_from_records(doc, data.get("s4_impact_summary", []))
        add_paragraph(doc, data.get("s4_reporting_impact_text", ""))

        add_heading(doc, "9. Gap Analysis Summary", 1)
        add_table_from_records(doc, data.get("gap_analysis_summary", []))
        add_paragraph(doc, data.get("key_gaps_text", ""))

        add_heading(doc, "10. Opportunity Areas", 1)
        add_table_from_records(doc, data.get("improvement_opportunity_summary", []))
        add_paragraph(doc, data.get("opportunity_areas_text", ""))

        add_heading(doc, "11. Business Value", 1)
        add_table_from_records(doc, data.get("potential_impact_summary", []))
        add_paragraph(doc, data.get("business_value_text", ""))

        add_heading(doc, "12. Recommended Next Steps", 1)
        add_table_from_records(doc, data.get("recommended_focus_areas", []))
        add_paragraph(doc, data.get("recommended_next_steps_text", ""))

        add_heading(doc, "13. Appendix A — Reporting Inventory", 1)
        add_table_from_records(doc, data.get("appendix_reporting_inventory", []))
        add_paragraph(doc, data.get("appendix_reporting_inventory_text", ""))

        add_heading(doc, "14. Appendix B — S/4 Reporting Impact Analysis", 1)
        add_table_from_records(doc, data.get("appendix_s4_impact_analysis", []))
        add_paragraph(doc, data.get("appendix_s4_impact_analysis_text", ""))

        add_heading(doc, "15. Appendix C — Reporting Overlap Analysis", 1)
        add_table_from_records(doc, data.get("appendix_reporting_overlap_analysis", []))
        add_paragraph(doc, data.get("appendix_reporting_overlap_analysis_text", ""))

        add_heading(doc, "16. Appendix D — Data Source Mapping", 1)
        add_table_from_records(doc, data.get("appendix_data_source_mapping", []))
        add_paragraph(doc, data.get("appendix_data_source_mapping_text", ""))

        add_heading(doc, "17. Appendix E — Critical Reports", 1)
        add_table_from_records(doc, data.get("appendix_critical_reports", []))
        add_paragraph(doc, data.get("appendix_critical_reports_text", ""))

        add_heading(doc, "Critical Report Summary", 2)
        add_paragraph(doc, data.get("critical_report_summary", ""))

        add_heading(doc, "18. Appendix F — Analytics Stakeholder Map", 1)
        add_table_from_records(doc, data.get("analytics_responsibility_model", []))
        add_table_from_records(doc, data.get("stakeholder_interview_summary", []))
        add_table_from_records(doc, data.get("responsibility_gaps", []))
        add_paragraph(doc, data.get("analytics_ownership_overview_text", ""))

    # --------------------
    # Analytics Modernization Roadmap
    # --------------------
    elif assessment_type == "Analytics Modernization Roadmap":

        add_heading(doc, "3. Modernization Drivers", 1)
        add_table_from_records(doc, data.get("modernization_drivers", []))

        add_heading(doc, "4. Current-State Architecture", 1)
        add_table_from_records(doc, data.get("current_state_architecture", []))

        add_heading(doc, "5. Future-State Architecture", 1)
        add_table_from_records(doc, data.get("future_state_architecture", []))

        add_heading(doc, "6. Capability Gap Summary", 1)
        add_table_from_records(doc, data.get("capability_gap_summary", []))

        add_heading(doc, "7. Platform Recommendations", 1)
        add_table_from_records(doc, data.get("platform_recommendations", []))

        add_heading(doc, "8. Workstream Plan", 1)
        add_table_from_records(doc, data.get("workstream_plan", []))

        add_heading(doc, "9. Risk Mitigation Plan", 1)
        add_table_from_records(doc, data.get("risk_mitigation_plan", []))

        add_heading(doc, "10. Investment Summary", 1)
        add_table_from_records(doc, data.get("investment_summary", []))

        add_heading(doc, "11. Business Value", 1)
        add_paragraph(doc, data.get("business_value_text", ""))
        add_table_from_records(doc, data.get("potential_impact_summary", []))

    # --------------------
    # AI Opportunity Assessment
    # --------------------
    elif assessment_type == "AI Opportunity Assessment":

        add_heading(doc, "3. Top AI Opportunities", 1)
        add_table_from_records(doc, data.get("top_ai_opportunities", []))

        add_heading(doc, "4. AI Use Case Inventory", 1)
        add_table_from_records(doc, data.get("ai_use_case_inventory", []))

        add_heading(doc, "5. Automation Candidates", 1)
        add_table_from_records(doc, data.get("automation_candidates", []))

        add_heading(doc, "6. Decision Support Opportunities", 1)
        add_table_from_records(doc, data.get("decision_support_opportunities", []))

        add_heading(doc, "7. Data Readiness Summary", 1)
        add_table_from_records(doc, data.get("data_readiness_summary", []))

        add_heading(doc, "8. AI Roadmap", 1)
        add_table_from_records(doc, data.get("ai_roadmap", []))

        add_heading(doc, "9. Risk and Governance Considerations", 1)
        add_table_from_records(doc, data.get("risk_and_governance_considerations", []))

        add_heading(doc, "10. Business Value", 1)
        add_paragraph(doc, data.get("business_value_text", ""))
        add_table_from_records(doc, data.get("potential_impact_summary", []))

        add_heading(doc, "11. Recommended Next Steps", 1)
        add_paragraph(doc, data.get("recommended_next_steps_text", ""))

    add_heading(doc, "Key Observations", 1)
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
def validate_output(data, assessment_type):
    if not data:
        return False

    framework = ASSESSMENT_FRAMEWORKS.get(assessment_type)

    if not framework:
        return False

    required_keys = []

    for section in framework["sections"]:
        required_keys.extend(section["keys"])

    for key in required_keys:
        value = data.get(key)

        if value is None:
            return False

        if isinstance(value, str):
            if not value.strip():
                return False
            if "to be validated" in value.lower():
                return False

        if isinstance(value, list):
            if len(value) == 0:
                return False
            if "to be validated" in str(value).lower():
                return False

        if isinstance(value, dict):
            if len(value) == 0:
                return False
            if "to be validated" in str(value).lower():
                return False

    return True

# --------------------
# Assessment Frameworks
# --------------------
ASSESSMENT_FRAMEWORKS = {
    "Analytics Gap Assessment": {
        "title": "Analytics Gap Assessment",
        "sections": [
            {
                "section_name": "Executive Overview",
                "keys": [
                    "engagement_overview_text",
                    "executive_summary_text",
                    "top_priorities",
                    "implementation_roadmap"
                ],
                "instructions": """
Focus on current-state analytics, reporting, governance, data ownership, system complexity, and decision-support gaps.

After each table, generate a 1-2 paragraph executive narrative summarizing:
- what the table shows
- why it matters
- operational implications
- business risks
- why leadership should care

top_priorities must be exactly 5 rows with:
Priority, Why It Matters, Business Impact, Time Horizon, Executive Owner

implementation_roadmap must be exactly 3 rows with:
Phase, Timeline, Key Actions, Business Outcome, Dependencies
"""
            },
            {
                "section_name": "Current State and Gap Analysis",
                "keys": [
                    "analytics_environment_snapshot",
                    "analytics_complexity_text",
                    "analytics_complexity_snapshot",
                    "gap_severity_heatmap",
                    "gap_observations_text",
                    "gap_analysis_summary",
                    "recommended_focus_areas"
                ],
                "instructions": """
Focus on current-state analytics, reporting, governance, data ownership, system complexity, and decision-support gaps.

Tables must use business-specific rows, not generic labels.
Each table row should explain:
Business Area, Current State, Where It Breaks, Business Impact, Why It Matters, Recommended Action, Priority

After each table, generate a 1-2 paragraph executive narrative summarizing:
- what the table shows
- why it matters
- operational implications
- business risks
- why leadership should care

Do not repeat the table. Interpret the findings for the customer.
"""
            },
            {
                "section_name": "Reporting, S/4 Impact, and Business Value",
                "keys": [
                    "current_landscape_text",
                    "current_architecture_summary",
                    "reporting_inventory_text",
                    "reporting_landscape_summary",
                    "s4_reporting_impact_text",
                    "s4_impact_summary",
                    "business_value_text",
                    "potential_impact_summary",
                    "recommended_next_steps_text"
                ],
                "instructions": """
Focus on reporting inventory, architecture, S/4HANA reporting impact where relevant, business value, and next steps.

After each table, generate a 1-2 paragraph executive narrative summarizing:
- what the table shows
- why it matters
- operational implications
- business risks
- why leadership should care
"""
            },
            {
                "section_name": "Appendices",
                "keys": [
                    "appendix_reporting_inventory",
                    "appendix_s4_impact_analysis",
                    "appendix_reporting_overlap_analysis",
                    "appendix_data_source_mapping",
                    "appendix_critical_reports",
                    "critical_report_summary",
                    "analytics_ownership_overview",
                    "analytics_responsibility_model",
                    "stakeholder_interview_summary",
                    "responsibility_gaps",
                    "key_observations_text"
                ],
                "instructions": """
Populate all appendices. No placeholders.used.used.

appendix_reporting_inventory columns:
Report Name, Business Function, Frequency, Current Owner, Current Issue, Recommended Disposition

appendix_s4_impact_analysis columns:
Process Area, Current Reporting Dependency, S/4HANA Impact, Risk Level, Required Action

appendix_reporting_overlap_analysis columns:
Report / Dashboard, Overlap Area, Duplicative Source, Business Risk, Recommended Action

appendix_data_source_mapping columns:
Data Source, Business Function, Current Usage, Integration Issue, Future-State Recommendation

appendix_critical_reports columns:
Critical Report, Executive Owner, Business Purpose, Risk If Unavailable, Modernization Priority
"""
            }
        ]
    },

    "Analytics Modernization Roadmap": {
        "title": "Analytics Modernization Roadmap",
        "sections": [
            {
                "section_name": "Executive Roadmap Overview",
                "keys": [
                    "engagement_overview_text",
                    "executive_summary_text",
                    "modernization_drivers",
                    "top_priorities",
                    "implementation_roadmap"
                ],
                "instructions": """
Focus on why modernization is needed, what needs to change, and the phased path forward.

modernization_drivers columns:
Driver, Current Constraint, Business Impact, Modernization Response, Priority

implementation_roadmap must include:
Phase, Timeline, Key Actions, Business Outcome, Dependencies
"""
            },
            {
                "section_name": "Current vs Future State",
                "keys": [
                    "current_state_architecture",
                    "future_state_architecture",
                    "capability_gap_summary",
                    "platform_recommendations"
                ],
                "instructions": """
Compare current-state and future-state analytics capabilities.

Tables should explain:
Capability Area, Current State, Future State, Gap, Recommended Action, Priority
"""
            },
            {
                "section_name": "Execution Plan",
                "keys": [
                    "workstream_plan",
                    "risk_mitigation_plan",
                    "investment_summary",
                    "business_value_text",
                    "potential_impact_summary"
                ],
                "instructions": """
Create a practical execution plan with workstreams, risks, dependencies, estimated value, and investment considerations.
"""
            }
        ]
    },

    "AI Opportunity Assessment": {
        "title": "AI Opportunity Assessment",
        "sections": [
            {
                "section_name": "AI Executive Opportunity Overview",
                "keys": [
                    "engagement_overview_text",
                    "executive_summary_text",
                    "top_ai_opportunities",
                    "implementation_roadmap"
                ],
                "instructions": """
Focus on where AI can create business value, reduce manual work, improve decisions, and accelerate operations.

top_ai_opportunities columns:
Use Case, Business Function, Current Pain Point, AI Opportunity, Business Value, Complexity, Priority
"""
            },
            {
                "section_name": "AI Use Case Portfolio",
                "keys": [
                    "ai_use_case_inventory",
                    "automation_candidates",
                    "decision_support_opportunities",
                    "data_readiness_summary"
                ],
                "instructions": """
Identify realistic AI use cases and readiness gaps.

Use case tables should include:
Use Case, Process Area, Required Data, Expected Benefit, Complexity, Recommended Next Step
"""
            },
            {
                "section_name": "AI Roadmap and Value",
                "keys": [
                    "ai_roadmap",
                    "risk_and_governance_considerations",
                    "business_value_text",
                    "potential_impact_summary",
                    "recommended_next_steps_text"
                ],
                "instructions": """
Create a phased AI roadmap with governance, risk, change management, data readiness, and measurable value.
"""
            }
        ]
    }
}

# --------------------
# Generate Button
# --------------------
if st.button("Generate Assessment Outputs", key="main_generate_btn"):

    if not client_name:
        st.warning("Enter a client name first.")
    else:
        file_content = read_uploaded_files(uploaded_files)
        company_research = research_company(client_name, industry)

        with st.spinner("Generating assessment content..."):
            max_retries = 3
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
        
                if validate_output(data, assessment_type):
                    break
                else:
                    st.warning(f"Regenerating output attempt {attempt + 1}: missing sections or placeholder text found.")
        
            if not validate_output(data, assessment_type):
                st.error("Failed to generate a complete assessment after retries.")
                data = {}

        st.session_state.assessment_data = data

        if data:
            with st.spinner("Creating Word document..."):
                st.session_state.word_doc = build_docx(data, client_name, assessment_type)

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
