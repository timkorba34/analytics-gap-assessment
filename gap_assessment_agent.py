import io
import json
import os
import streamlit as st
from openai import OpenAI
from docx import Document
import PyPDF2
import pandas as pd


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
    type=["txt", "csv", "pdf"],
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

        if file.type == "text/plain":
            content += file.read().decode("utf-8", errors="ignore")

        elif file.type == "text/csv":
            df = pd.read_csv(file)
            content += df.to_string(index=False)

        elif file.type == "application/pdf":
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                content += page.extract_text() or ""

    return content


# --------------------
# Generate Assessment JSON
# --------------------
def generate_assessment_json(client_name, industry, assessment_type, notes, file_content):
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

Appendix sections are REQUIRED and must be populated using reasonable inferred values from discovery notes.

Minimum rows required:

appendix_reporting_inventory = minimum 8 rows
appendix_reporting_overlap_analysis = minimum 5 rows
appendix_data_source_mapping = minimum 8 rows
appendix_critical_reports = minimum 6 rows
analytics_ownership_overview = minimum 6 rows
stakeholder_interview_summary = minimum 6 rows

Only use "To be validated" in one specific field when necessary, never for entire section.

Required table structures:

engagement_scope_summary:
[
  {{"Area": "...", "In Scope": "...", "Assessment Focus": "..."}}
]

analytics_environment_snapshot:
[
  {{"Area": "POS", "Current Tool / Process": "...", "Observed Issue": "...", "Business Impact": "..."}}
]

analytics_complexity_snapshot:
[
  {{"Complexity Driver": "...", "Current Observation": "...", "Impact": "...", "Severity": "..."}}
]

gap_severity_heatmap:
[
  {{"Gap Domain": "...", "Current State": "...", "Severity": "High/Medium/Low", "Business Impact": "...", "Recommended Action": "..."}}
]

current_architecture_summary:
[
  {{"Layer": "...", "Current State": "...", "Gap / Limitation": "...", "Future-State Consideration": "..."}}
]

reporting_landscape_summary:
[
  {{"Reporting Area": "...", "Current Method": "...", "Issue": "...", "Improvement Opportunity": "..."}}
]

s4_impact_summary:
[
  {{"Area": "...", "Impact": "Not applicable unless SAP/S4 is identified", "Validation Needed": "..."}}
]

gap_analysis_summary:
[
  {{"Gap": "...", "Evidence from Discovery": "...", "Business Impact": "...", "Priority": "..."}}
]

improvement_opportunity_summary:
[
  {{"Opportunity": "...", "Description": "...", "Expected Value": "...", "Priority": "..."}}
]

potential_impact_summary:
[
  {{"Value Lever": "...", "Potential Impact": "...", "How It Helps": "..."}}
]

recommended_focus_areas:
[
  {{"Focus Area": "...", "Why It Matters": "...", "Recommended Next Step": "..."}}
]
appendix_reporting_inventory:
[
 {{"Report Name":"Daily Sales by Store","Business Owner":"Operations","Current Tool":"Excel","Frequency":"Daily","Pain Point":"Manual consolidation","Priority":"High"}}
]

appendix_reporting_overlap_analysis:
[
 {{"Metric":"Sales by Store","Found In":"POS export / Finance file / Manager tracker","Issue":"Multiple versions of truth","Recommendation":"Centralize metric ownership"}}
]

appendix_data_source_mapping:
[
 {{"Source System":"Toast POS","Data Type":"Transactions","Refresh Frequency":"Daily","Current Method":"CSV Export","Future Method":"API Integration"}}
]

appendix_critical_reports:
[
 {{"Report":"Daily Flash Sales","Business Impact":"High","Consumers":"CEO / Ops","Current State":"Manual","Future State":"Automated Dashboard"}}
]

analytics_ownership_overview:
[
 {{"Domain":"Sales Reporting","Current Owner":"Regional Managers","Future Owner":"Analytics COE","Gap":"No formal ownership"}}
]

stakeholder_interview_summary:
[
 {{"Stakeholder":"Owner","Priority":"Growth visibility","Pain Point":"No single dashboard","Need":"Daily KPIs"}}
]


Rules:
- Return ONLY valid JSON. No markdown.
- Every table field must be an array of objects.
- Do not return nested objects inside table fields.
- Do not return Python-style lists as strings.
- If data is unavailable, still create 3-5 reasonable assessment rows using "To be validated" only in the Evidence / Validation Needed column.
- Do not include S/4HANA content unless SAP, ECC, or S/4HANA is mentioned in the notes.
- For non-SAP clients, set S/4HANA sections to "Not applicable based on current discovery inputs."
- Avoid generic consulting language. Tie every gap and recommendation to the client facts.
- Use business-friendly language for executives.
- All narrative text fields must be detailed and presentation-ready.
- Each narrative section should contain 2-3 strong paragraphs.
- Narratives must explain:
  1. Current state observations
  2. Why the issue matters
  3. Business risks / impacts
  4. Likely root causes
  5. Future-state opportunity
  6. Recommended path forward

- Avoid generic filler language.
- Tie commentary to the client discovery notes.
- Write like a senior consulting advisor preparing material for executives.
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "Return only valid JSON. No markdown."},
            {"role": "user", "content": prompt}
        ],
        temperature=0.2
    )

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


# --------------------
# Generate Button
# --------------------
if st.button("Generate Word Assessment"):
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

        with st.spinner("Creating Word document..."):
            docx_file = build_docx(data, client_name)

        safe_client_name = client_name.replace(" ", "_").replace("'", "").replace("/", "_")

        st.success("Assessment generated.")

        st.download_button(
            label="Download Word Assessment",
            data=docx_file.getvalue(),
            file_name=f"{safe_client_name}_Gap_Assessment.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
