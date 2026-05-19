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
from docx import Document
from tavily import TavilyClient
from graphviz import Digraph


# --------------------
# Page Setup
# --------------------
st.set_page_config(page_title="AI Gap Assessment Builder", layout="wide")

import base64

def get_base64(file_path):
    with open(file_path, "rb") as f:
        return base64.b64encode(f.read()).decode()

bg_image = get_base64("copilot_bg.png")
robot_image = get_base64("robot_bg.png")

st.markdown(
    f"""
    <style>

    .stApp {{
        background: linear-gradient(rgba(5,10,25,0.96), rgba(5,10,25,0.98));
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
        color: white;
    }}

    .main-title {{
        font-size: 54px;
        font-weight: 700;
        color: #ffffff !important;
        margin-bottom: 10px;
        line-height: 1.15;
    }}

    .sub-title {{
        font-size: 22px;
        color: #c7d2fe !important;
        margin-bottom: 20px;
        line-height: 1.5;
    }}

    .glass {{
        background: rgba(15, 23, 42, 0.72);
        backdrop-filter: blur(12px);
        padding: 30px;
        border-radius: 20px;
        border: 1px solid rgba(255,255,255,0.12);
        box-shadow: 0 0 30px rgba(0,0,0,0.4);
        margin-bottom: 25px;
    }}

    .hero-container {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: linear-gradient(
            90deg,
            rgba(10,15,40,0.95),
            rgba(15,23,42,0.82)
        );
        border-radius: 28px;
        padding: 50px;
        overflow: hidden;
        margin-bottom: 30px;
        border: 1px solid rgba(255,255,255,0.08);
    }}

    .hero-left {{
        width: 58%;
    }}

    .hero-right {{
        width: 38%;
        text-align: right;
    }}

    .hero-right img {{
        width: 100%;
        max-width: 420px;
    }}

    .hero-title {{
        font-size: 64px;
        font-weight: 800;
        color: white;
        line-height: 1.1;
        margin-bottom: 20px;
    }}

    .hero-subtitle {{
        font-size: 22px;
        color: #c7d2fe;
        line-height: 1.6;
    }}

    .feature-card {{
        background: rgba(20, 30, 60, 0.72);
        padding: 24px;
        border-radius: 18px;
        border: 1px solid rgba(255,255,255,0.12);
        margin-bottom: 20px;
        min-height: 160px;
    }}

    .feature-card h3 {{
        color: #ffffff !important;
        font-size: 26px;
        margin-bottom: 15px;
    }}

    .feature-card p {{
        color: #e5e7eb !important;
        font-size: 17px;
        line-height: 1.5;
    }}

    label,
    .stMarkdown p,
    .stMarkdown span {{
        color: #ffffff !important;
    }}

    label[data-testid="stWidgetLabel"] {{
        color: #ffffff !important;
        font-weight: 600 !important;
    }}

    .stTextInput input,
    .stTextArea textarea {{
        background-color: rgba(15, 23, 42, 0.85) !important;
        color: #ffffff !important;
        border: 1px solid rgba(255,255,255,0.25) !important;
        border-radius: 10px !important;
    }}

    .stSelectbox div[data-baseweb="select"] > div {{
        background-color: rgba(15, 23, 42, 0.85) !important;
        color: #ffffff !important;
        border: 1px solid rgba(255,255,255,0.25) !important;
        border-radius: 10px !important;
    }}

   /* Uploaded file pill background */
div[data-testid="stFileUploaderFile"] {{
    background: #111a33 !important;
    border: 1px solid rgba(120, 160, 255, 0.6) !important;
    border-radius: 12px !important;
    }}
    
    /* Inner file pill containers */
    div[data-testid="stFileUploaderFile"] > div {{
        background: #111a33 !important;
        color: #ffffff !important;
    }}
    
    /* Filename text */
    div[data-testid="stFileUploaderFile"] span,
    div[data-testid="stFileUploaderFile"] p,
    div[data-testid="stFileUploaderFile"] small {{
        color: #ffffff !important;
    }}
    
    /* File icon box */
    div[data-testid="stFileUploaderFile"] svg {{
        color: #ffffff !important;
        fill: #ffffff !important;
    }}
    
    /* Remove button */
    div[data-testid="stFileUploaderFile"] button {{
        background: #0b1f44 !important;
        color: #ffffff !important;
        border: 1px solid rgba(120, 160, 255, 0.6) !important;
    }}

    .stFileUploader * {{
        color: #ffffff !important;
    }}

    div.stButton > button {{
        background: linear-gradient(90deg, #2563eb, #7c3aed);
        color: white !important;
        border-radius: 14px;
        height: 55px;
        font-size: 18px;
        font-weight: 600;
        border: none;
        width: 100%;
    }}

    div.stButton > button:hover {{
        background: linear-gradient(90deg, #1d4ed8, #6d28d9);
        color: white !important;
    }}

    .stFileUploader label div {{
    color: white !important;
    }}

    .stFileUploader section {{
        background-color: rgba(10, 20, 45, 0.85) !important;
        border: 1px dashed rgba(255,255,255,0.3) !important;
        border-radius: 12px;
    }}

    .stFileUploader section:hover {{
        border: 1px solid #4da6ff !important;
    }}

    .stFileUploader div[data-testid="stFileUploaderDropzone"] {{
        background-color: rgba(10,20,45,0.9) !important;
        color: white !important;
    }}

    .stFileUploader button {{
        background-color: #0b1f44 !important;
        color: white !important;
        border: 1px solid rgba(255,255,255,0.2) !important;
        border-radius: 8px !important;
    }}

    .stFileUploader button:hover {{
        border: 1px solid #4da6ff !important;
        color: #4da6ff !important;
    }}

    </style>
    """,
    unsafe_allow_html=True
)

st.markdown(
    f"""<div class="hero-container">
<div class="hero-left">
<div class="hero-title">Analytics Modernization Assessment Copilot</div>
<div class="hero-subtitle">
AI-powered executive analytics assessments, modernization roadmaps,
S/4HANA readiness analysis, and actionable remediation planning.
</div>
</div>
<div class="hero-right">
<img src="data:image/png;base64,{robot_image}">
</div>
</div>""",
    unsafe_allow_html=True
)

col1, col2, col3 = st.columns(3)

with col1:
    st.markdown("""
    <div class='feature-card'>
        <h3>AI-Driven Insights</h3>
        <p>Identify analytics gaps, risks, and reporting challenges automatically.</p>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class='feature-card'>
        <h3>Executive Ready</h3>
        <p>Generate boardroom-quality Word documents and PowerPoint presentations.</p>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown("""
    <div class='feature-card'>
        <h3>S/4HANA Readiness</h3>
        <p>Assess reporting impacts, governance gaps, and modernization priorities.</p>
    </div>
    """, unsafe_allow_html=True)

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
# Company Search - Drop Down
# --------------------
def search_companies(query):
    try:
        response = tavily_client.search(
            query=f"{query} official company name official website",
            search_depth="basic",
            max_results=5,
            exclude_domains=[
                "linkedin.com",
                "facebook.com",
                "twitter.com",
                "x.com",
                "glassdoor.com",
                "zoominfo.com",
                "seamless.ai",
                "rocketreach.co",
                "wikipedia.org"
            ]
        )

        search_context = ""

        for result in response.get("results", []):
            search_context += f"""
Title: {result.get("title", "")}
URL: {result.get("url", "")}
Content: {result.get("content", "")}
"""

        prompt = f"""
You are identifying the correct official company name.

User typed:
{query}

Search results:
{search_context}

Return only the official company name.
Do not return page titles.
Do not return descriptions.
Do not return URLs.
Do not return LinkedIn, ZoomInfo, Seamless, employee names, or marketing text.
Return one clean company name only.

Example:
User typed: PIM Brands
Return: PIM Brands, Inc.
"""

        result = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Return only one clean official company name."},
                {"role": "user", "content": prompt}
            ],
            temperature=0
        )

        company_name = result.choices[0].message.content.strip()

        return [company_name]

    except Exception:
        return [query.strip()]
# --------------------
# UI Inputs
# --------------------
st.markdown("<div class='glass'>", unsafe_allow_html=True)

client_search = st.text_input("Search Client")

company_options = []

if client_search:
    company_options = search_companies(client_search)

if company_options:

    client_name = st.selectbox(
        "Select Company",
        company_options
    )

else:
    client_name = client_search
    
industry = st.selectbox(
    "Industry",
    [
        "Select an industry",
        "Accounting",
        "Airlines and Aviation",
        "Alternative Medicine",
        "Animation and Post-Production",
        "Apparel Manufacturing",
        "Architecture and Planning",
        "Automotive",
        "Banking",
        "Biotechnology Research",
        "Broadcast Media Production and Distribution",
        "Building Materials",
        "Business Consulting and Services",
        "Chemical Manufacturing",
        "Civil Engineering",
        "Computer and Network Security",
        "Construction",
        "Consumer Services",
        "Data Infrastructure and Analytics",
        "Defense and Space Manufacturing",
        "Education",
        "Education Administration Programs",
        "Electric Power Generation",
        "Entertainment Providers",
        "Environmental Services",
        "Events Services",
        "Executive Offices",
        "Facilities Services",
        "Farming",
        "Financial Services",
        "Food and Beverage Manufacturing",
        "Government Administration",
        "Healthcare",
        "Hospitals and Health Care",
        "Hospitality",
        "Human Resources Services",
        "Industrial Machinery Manufacturing",
        "Information Services",
        "Information Technology & Services",
        "Insurance",
        "Investment Banking",
        "IT Services and IT Consulting",
        "Law Practice",
        "Legal Services",
        "Life Sciences",
        "Logistics and Supply Chain",
        "Machinery Manufacturing",
        "Manufacturing",
        "Market Research",
        "Medical Equipment Manufacturing",
        "Mining",
        "Non-profit Organizations",
        "Oil and Gas",
        "Pharmaceutical Manufacturing",
        "Primary and Secondary Education",
        "Professional Services",
        "Public Relations and Communications Services",
        "Real Estate",
        "Research Services",
        "Retail",
        "Retail Apparel and Fashion",
        "Semiconductor Manufacturing",
        "Software Development",
        "Sports",
        "Telecommunications",
        "Transportation/Trucking/Railroad",
        "Travel Arrangements",
        "Utilities",
        "Warehousing and Storage",
        "Wholesale",
        "Wireless Services"
        "Other"
    ]
)

source_system = st.selectbox(
    "Source System",
    [
        "SAP ECC",
        "SAP S/4HANA",
        "Oracle EBS",
        "Oracle Fusion",
        "JD Edwards",
        "Microsoft Dynamics",
        "Infor",
        "NetSuite",
        "Multiple ERP Systems",
        "Other"
    ]
)

migration_type = st.selectbox(
    "Migration Type",
    [
        "Brownfield",
        "Greenfield",
        "Smartfield / Selective Data Transition",
        "Not Yet Defined"
    ]
)

target_source_system = st.selectbox(
    "Target Source System",
    [
        "SAP S/4HANA",
        "Oracle Fusion",
        "SAP ECC (No Migration)",
        "Microsoft Dynamics",
        "Multiple ERP Systems",
        "Other"
    ]
)

assessment_type = st.selectbox(
    "Assessment Type",
    [
        "Analytics Gap Assessment",
        "Analytics Modernization Assessment",
        "S/4HANA Impact Assessment",
        "AI Readiness Assessment",
        "Data Governance Assessment",
        "Reporting Rationalization Assessment"
    ]
)

uploaded_files = st.file_uploader(
    "Upload Discovery Notes / Supporting Files",
    type=["txt", "csv", "pdf", "xls", "xlsx", "doc", "docx"],
    accept_multiple_files=True
)



notes = st.text_area("Paste Additional Notes", height=250)

safe_client_name = client_name.strip().replace(" ", "_") if client_name else "Client"

st.markdown("</div>", unsafe_allow_html=True)

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

            elif file_name.endswith(".docx"):

                doc = Document(file)

                for paragraph in doc.paragraphs:
                    content += paragraph.text + "\n"

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
Source System: {source_system}
Migration Type: {migration_type}
Target Platform: {target_platform}
Assessment Type: {assessment_type}

Adjust recommendations, impacted reports, dependencies,
table mappings, architecture, and roadmap based on these selections.

Transformation Context:
Current Source System: {source_system}
Target Source System: {target_source_system}
Migration Type: {migration_type}

The assessment must explicitly reference this transformation path throughout the report. 
Do not generically say “S/4HANA migration.” Explain how the selected migration type changes

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

If the required keys contain "transformation_context", generate it as an array.

Format:

"transformation_context": [
  {{
    "area": "ERP",
    "current_state": "SAP ECC",
    "target_state": "SAP S/4HANA",
    "migration_impact": "Brownfield migration retains existing configurations and customizations, requiring validation of existing reporting dependencies and custom code."
  }},
  {{
    "area": "Finance",
    "current_state": "BKPF/BSEG/FAGLFLEXA",
    "target_state": "ACDOCA",
    "migration_impact": "Universal Journal consolidation impacts financial reporting logic and report structures."
  }},
  {{
    "area": "Inventory",
    "current_state": "MKPF/MSEG",
    "target_state": "MATDOC",
    "migration_impact": "Inventory reporting logic and historical movement analysis require validation."
  }},
  {{
    "area": "Sales",
    "current_state": "VBAK/VBAP/VBRK/VBRP",
    "target_state": "S/4HANA CDS Views",
    "migration_impact": "Order-to-cash reporting dependencies and KPIs require revalidation."
  }},
  {{
    "area": "Reporting",
    "current_state": "BusinessObjects / Excel / BW",
    "target_state": "S/4HANA Embedded Analytics / Datasphere",
    "migration_impact": "Reports should be categorized into retain, remediate, rebuild, or retire."
  }}
]

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
        doc.add_paragraph("No records were generated for this section.")
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
    doc.add_paragraph("")
    add_paragraph(doc, data.get("engagement_overview_text", ""))

    add_heading(doc, "2. Executive Summary", 1)
    doc.add_paragraph("")
    add_paragraph(doc, data.get("executive_summary_text", ""))

    if data.get("top_priorities"):
        add_heading(doc, "Executive Priorities", 2)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("top_priorities", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("top_priorities_text", ""))  

    # --------------------
    # Analytics Gap Assessment
    # --------------------
    if assessment_type == "Analytics Gap Assessment":

        add_heading(doc, "3. S/4HANA Transformation Context", 1)
        doc.add_paragraph("")
        add_table_from_records(
        doc,
        data.get("transformation_context", []))
        doc.add_paragraph("")
        add_paragraph(
        doc,
        data.get("transformation_context_summary", ""))

        add_heading(doc, "4. Current Analytics Ecosystem", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("current_system_inventory", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("current_data_flow_summary", ""))

        add_heading(doc, "5. Reporting Dependency Map", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("reporting_dependency_map", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("architecture_risk_summary", ""))

        add_heading(doc, "6. Analytics Complexity and Operational Risk", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("analytics_complexity_snapshot", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("analytics_complexity_text", ""))

        add_heading(doc, "7. Gap Severity Heatmap", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("gap_severity_heatmap", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("gap_observations_text", ""))

        add_heading(doc, "8. Reporting Inventory Summary", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("reporting_landscape_summary", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("reporting_inventory_text", ""))

        add_heading(doc, "9. S/4HANA Reporting Impact", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("s4_impact_summary", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("s4_reporting_impact_text", ""))

        add_heading(doc, "10. Gap Analysis Summary", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("gap_analysis_summary", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("key_gaps_text", ""))

        add_heading(doc, "11. Opportunity Areas", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("improvement_opportunity_summary", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("opportunity_areas_text", ""))

        add_heading(doc, "12. Business Value", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("potential_impact_summary", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("business_value_text", ""))

        add_heading(doc, "13. Gap Remediation Roadmap", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("s4_analytics_roadmap", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("s4_analytics_roadmap_text", ""))

        add_heading(doc, "14. Recommended Remediation Actions and Execution Plan", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("recommended_focus_areas", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("recommended_next_steps_text", ""))

        add_heading(doc, "15. Appendix A — Reporting Inventory", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("appendix_reporting_inventory", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("appendix_reporting_inventory_text", ""))

        add_heading(doc, "16. Appendix B — S/4 Reporting Impact Analysis", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("appendix_s4_impact_analysis", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("appendix_s4_impact_analysis_text", ""))

        add_heading(doc, "17. Appendix C — Reporting Overlap Analysis", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("appendix_reporting_overlap_analysis", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("appendix_reporting_overlap_analysis_text", ""))

        add_heading(doc, "18. Appendix D — Data Source Mapping", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("appendix_data_source_mapping", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("appendix_data_source_mapping_text", ""))

        add_heading(doc, "19. Appendix E — Critical Reports", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("appendix_critical_reports", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("appendix_critical_reports_text", ""))

        add_heading(doc, "Critical Report Summary", 2)
        doc.add_paragraph("")
        add_paragraph(doc, data.get("critical_report_summary", ""))

        add_heading(doc, "20. Appendix F — Analytics Stakeholder Map", 1)
        doc.add_paragraph("")
        add_table_from_records(doc, data.get("analytics_responsibility_model", []))
        add_table_from_records(doc, data.get("stakeholder_interview_summary", []))
        add_table_from_records(doc, data.get("responsibility_gaps", []))
        doc.add_paragraph("")
        add_paragraph(doc, data.get("analytics_ownership_overview_text", ""))

    # --------------------
    # Analytics Modernization Roadmap
    # --------------------
    elif assessment_type == "Analytics Modernization Roadmap":

        add_heading(doc, "3. Modernization Drivers", 1)
        add_table_from_records(doc, data.get("modernization_drivers", []))
        doc.add_paragraph("")

        add_heading(doc, "4. Current-State Architecture", 1)
        add_table_from_records(doc, data.get("current_state_architecture", []))
        doc.add_paragraph("")

        add_heading(doc, "5. Future-State Architecture", 1)
        add_table_from_records(doc, data.get("future_state_architecture", []))
        doc.add_paragraph("")

        add_heading(doc, "6. Capability Gap Summary", 1)
        add_table_from_records(doc, data.get("capability_gap_summary", []))
        doc.add_paragraph("")

        add_heading(doc, "7. Platform Recommendations", 1)
        add_table_from_records(doc, data.get("platform_recommendations", []))
        doc.add_paragraph("")

        add_heading(doc, "8. Workstream Plan", 1)
        add_table_from_records(doc, data.get("workstream_plan", []))
        doc.add_paragraph("")

        add_heading(doc, "9. Risk Mitigation Plan", 1)
        add_table_from_records(doc, data.get("risk_mitigation_plan", []))
        doc.add_paragraph("")

        add_heading(doc, "10. Investment Summary", 1)
        add_table_from_records(doc, data.get("investment_summary", []))
        doc.add_paragraph("")

        add_heading(doc, "11. Business Value", 1)
        add_paragraph(doc, data.get("business_value_text", ""))
        add_table_from_records(doc, data.get("potential_impact_summary", []))
        doc.add_paragraph("")

    # --------------------
    # AI Opportunity Assessment
    # --------------------
    elif assessment_type == "AI Opportunity Assessment":

        add_heading(doc, "3. Top AI Opportunities", 1)
        add_table_from_records(doc, data.get("top_ai_opportunities", []))
        doc.add_paragraph("")

        add_heading(doc, "4. AI Use Case Inventory", 1)
        add_table_from_records(doc, data.get("ai_use_case_inventory", []))
        doc.add_paragraph("")

        add_heading(doc, "5. Automation Candidates", 1)
        add_table_from_records(doc, data.get("automation_candidates", []))
        doc.add_paragraph("")

        add_heading(doc, "6. Decision Support Opportunities", 1)
        add_table_from_records(doc, data.get("decision_support_opportunities", []))
        doc.add_paragraph("")

        add_heading(doc, "7. Data Readiness Summary", 1)
        add_table_from_records(doc, data.get("data_readiness_summary", []))
        doc.add_paragraph("")

        add_heading(doc, "8. AI Roadmap", 1)
        add_table_from_records(doc, data.get("ai_roadmap", []))
        doc.add_paragraph("")

        add_heading(doc, "9. Risk and Governance Considerations", 1)
        add_table_from_records(doc, data.get("risk_and_governance_considerations", []))
        doc.add_paragraph("")

        add_heading(doc, "10. Business Value", 1)
        add_paragraph(doc, data.get("business_value_text", ""))
        add_table_from_records(doc, data.get("potential_impact_summary", []))
        doc.add_paragraph("")

        add_heading(doc, "11. Recommended Next Steps", 1)
        add_paragraph(doc, data.get("recommended_next_steps_text", ""))
        doc.add_paragraph("")

        add_heading(doc, "Key Observations", 1)
        add_paragraph(doc, data.get("key_observations_text", ""))
        doc.add_paragraph("")

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
            "top_priorities_text",
            "s4_analytics_roadmap",
            "s4_analytics_roadmap_text"
        ],
        "instructions": """
top_priorities must be a table array with exactly 5 rows:
Priority, Why It Matters, Business Impact, Time Horizon, Executive Owner

s4_analytics_roadmap must be a post-assessment gap remediation roadmap.

Do not include an Assessment phase. The assessment has already been completed.

The roadmap must explain how the client should remediate the identified analytics, reporting, governance, data ownership, KPI, historical reporting, integration, and S/4HANA readiness gaps.

Each phase must be action-oriented and tied directly to fixing gaps identified in this report.

Use phases such as:
1. Stabilize Critical Reporting Risk
2. Define Data Ownership and KPI Governance
3. Remediate S/4HANA Reporting Dependencies
4. Modernize Priority Reports and Data Flows
5. Operationalize Analytics Governance and Continuous Improvement

Columns:
Phase, Timeline, Gap Addressed, Remediation Actions, Expected Outcome, Business Value, Dependencies
"""
    },
    {
        "section_name": "Current State and Gap Analysis",
        "keys": [
            "analytics_environment_snapshot",
            "analytics_environment_summary",
            "transformation_context",
            "transformation_context_summary",
            "current_system_inventory",
            "current_data_flow_summary",
            "reporting_dependency_map",
            "architecture_risk_summary",
            "analytics_complexity_text",
            "analytics_complexity_snapshot",
            "gap_severity_heatmap",
            "gap_observations_text",
            "gap_analysis_summary",
            "key_gaps_text",
            "recommended_focus_areas",
            "recommended_next_steps_text"
        ],
        "instructions": """

transformation_context must be a table array.

Columns:
Area,
Current State,
Target State,
Migration Impact

transformation_context_summary must be a 1-2 paragraph executive narrative explaining:

- current source-to-target migration path
- Brownfield/Greenfield impact
- report breakage risks
- source table impacts
- business process impacts
- remediation considerations

analytics_environment_snapshot must be a table array.
analytics_complexity_snapshot must be a table array.
gap_severity_heatmap must be a table array.
gap_analysis_summary must be a table array.

recommended_focus_areas must be a highly actionable post-assessment execution plan tied directly to the identified gaps in the report.

Do not generate generic consulting recommendations.

Every recommendation must:
- align to a specific identified gap
- explain the operational problem being solved
- identify the required business or technical action
- identify likely ownership
- explain why the action matters
- support movement toward execution readiness

Recommendations should reflect realistic follow-on activities that occur immediately after an assessment.

Include practical actions such as:
- validating critical reporting dependencies
- prioritizing report remediation
- defining KPI ownership
- establishing governance structures
- identifying high-risk S/4HANA reporting impacts
- aligning business stakeholders
- creating remediation workstreams
- defining implementation sequencing
- preparing roadmap funding
- creating implementation-ready scope
- defining follow-on SOW activities
- identifying quick wins
- preparing architecture decisions
- validating integration requirements
- creating a phased modernization strategy

Columns:
Recommendation Category,
Gap Addressed,
Recommended Action,
Business Outcome,
Priority,
Suggested Owner,
Execution Horizon,
Potential Follow-On Deliverable

All *_text and *_summary keys must be 1-2 paragraph narratives.

current_system_inventory must be a table array.

Columns:
Business Area,
Current Source System,
Reporting Tool,
Data Owner,
Integration Method,
Refresh Frequency,
Key Dependency,
Current Pain Point,
S/4HANA Risk

reporting_dependency_map must be a table array.

Columns:
Report / Dashboard,
Primary Source System,
Dependent Systems,
Business Function,
Criticality,
Current Risk,
S/4HANA Impact

current_data_flow_summary must be a 1-2 paragraph executive narrative explaining:
- how data flows today
- where fragmentation exists
- where manual effort exists
- where reporting bottlenecks occur
- operational risks
- S/4HANA disruption concerns

architecture_risk_summary must be a 1-2 paragraph executive narrative explaining:
- architecture weaknesses
- integration concerns
- reporting dependencies
- scalability limitations
- business continuity risks

recommended_next_steps_text must read like an executive transition plan from assessment into execution.

The narrative should explain:
- what leadership should do immediately
- what risks require urgent remediation
- where governance decisions are required
- which workstreams should begin first
- where implementation funding and alignment are needed
- what follow-on implementation activities should occur
- how the organization should move from assessment into execution

The tone should feel practical, operational, and implementation-oriented — not theoretical.
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
            "opportunity_areas_text",
            "improvement_opportunity_summary",
            "business_value_text",
            "potential_impact_summary"
        ],
        "instructions": """
current_architecture_summary must be a table array.
reporting_landscape_summary must be a table array.
s4_impact_summary must be a table array.
improvement_opportunity_summary must be a table array.
potential_impact_summary must be a table array.

All *_text keys must be 1-2 paragraph narratives.
"""
    },
    {
        "section_name": "Appendices",
        "keys": [
            "appendix_reporting_inventory",
            "appendix_reporting_inventory_text",
            "appendix_s4_impact_analysis",
            "appendix_s4_impact_analysis_text",
            "appendix_reporting_overlap_analysis",
            "appendix_reporting_overlap_analysis_text",
            "appendix_data_source_mapping",
            "appendix_data_source_mapping_text",
            "appendix_critical_reports",
            "appendix_critical_reports_text",
            "critical_report_summary",
            "analytics_ownership_overview_text",
            "analytics_responsibility_model",
            "stakeholder_interview_summary",
            "responsibility_gaps",
            "key_observations_text"
        ],
        "instructions": """
appendix_reporting_inventory must be a table array.
appendix_s4_impact_analysis must be a table array.
appendix_reporting_overlap_analysis must be a table array.
appendix_data_source_mapping must be a table array.
appendix_critical_reports must be a table array.
analytics_responsibility_model must be a table array.
stakeholder_interview_summary must be a table array.
responsibility_gaps must be a table array.

All *_text and *_summary keys must be 1-2 paragraph narratives.
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
                    "s4_analytics_roadmap"
                ],
                "instructions": """
Focus on why modernization is needed, what needs to change, and the phased path forward.

modernization_drivers columns:
Driver, Current Constraint, Business Impact, Modernization Response, Priority

s4_analytics_roadmap must represent the recommended phased execution plan AFTER the assessment.

The roadmap should focus on implementing the recommended analytics, reporting, governance, and S/4HANA readiness improvements.

Columns:
Phase, Timeline, Strategic Objective, Key Activities, Expected Outcome, Business Value, Dependencies
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
                    "s4_analytics_roadmap"
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
