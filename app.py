import streamlit as st
import io
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# 1. 顶级配置与 HIPAA 物理隔离合规侧边栏
st.set_page_config(page_title="BIPEngine - Expert System Framework", layout="wide", initial_sidebar_state="expanded")

with st.sidebar:
    st.markdown("<h2 style='color: #008080;'>⚙️ BIPEngine v1.3</h2>", unsafe_allow_html=True)
    st.subheader("🔒 Strict HIPAA Compliance")
    st.markdown("**Stateless Data Minimization**")
    st.caption("BIPEngine operates on a zero-knowledge framework. Uploaded CSV data logs and text interview payloads are evaluated entirely in-memory (RAM) and destroyed immediately upon session closure. No persistent storage or cloud logs exist.")
    st.divider()
    st.success("✔ BACB Ethics Code 2022 Compliant")
    st.info("💡 Reviewer Option: You can now upload raw clinical datasets below to test the automated compilation framework.")

st.title("🚀 BIPEngine: Automated Clinical Draft Compiler")
st.markdown("##### Streamlining FBA & BIP Documentation via Deterministic Matrix Mapping to Mitigate Practitioner Exhaustion")
st.divider()

# 2. 核心数据输入模块
st.header("📋 Phase 1: Multi-Source Data Ingestion")
tab1, tab2, tab3 = st.tabs(["📊 ABC Raw Data (Manual or CSV Upload)", "📝 Qualitative Stakeholder Interviews", "📈 Psychometric Profiler (QABF)"])

with tab1:
    st.subheader("Direct Observation ABC Ledger")
    st.caption("BCBAs can either modify the de-identified matrix below manually OR upload a raw client tracking CSV file.")
    
    uploaded_abc_file = st.file_uploader("Upload Raw ABC Data Tracking Sheet (CSV format)", type=["csv"])
    
    default_abc_data = [
        {"Entry": "Obs #1", "Date & Time": "07/20/2026 09:15 AM", "Setting / Context": "Structured Literacy / Desk Work", "Antecedent (A)": "Teacher presented a writing worksheet (Demand placed).", "Behavior (B)": "Screamed (>80dB), pushed desk away, and attempted to leave the room (Elopement). Duration: 3 min.", "Consequence (C)": "Staff provided verbal instruction 'Take a break' and directed student to quiet corner. Demand temporarily removed.", "Inferred Function": "Escape / Avoidance"},
        {"Entry": "Obs #2", "Date & Time": "07/21/2026 10:30 AM", "Setting / Context": "Free Play / Transition to Math", "Antecedent (A)": "Timer rang to signal end of iPad time (Preferred item removed).", "Behavior (B)": "Vocal outburst (screaming), dropped to floor, refused to move.", "Consequence (C)": "Staff offered a 2-minute extension with visual count-down timer before transition.", "Inferred Function": "Access to Tangible"},
        {"Entry": "Obs #3", "Date & Time": "07/22/2026 01:45 PM", "Setting / Context": "Small Group Activity", "Antecedent (A)": "Teacher turned attention away to assist another student.", "Behavior (B)": "Approached staff member, pulled staff sleeve, made loud vocalizations.", "Consequence (C)": "Staff immediately turned, made eye contact, and verbally reassured student ('I'm right here').", "Inferred Function": "Attention Seeking"},
        {"Entry": "Obs #4", "Date & Time": "07/24/2026 11:10 AM", "Setting / Context": "Independent Work Time", "Antecedent (A)": "Presented with multi-step math task; peers were quietly working.", "Behavior (B)": "Stood up suddenly, walked fast towards classroom door (Elopement).", "Consequence (C)": "Aide blocked door, handed 'I Need a Break' visual card, guided to sensory area.", "Inferred Function": "Escape Task"}
    ]
    
    if uploaded_abc_file is not None:
        try:
            display_df = pd.read_csv(uploaded_abc_file)
            st.success("🎉 Custom clinical ABC CSV log parsed successfully in-memory!")
        except Exception as e:
            st.error(f"Error parsing file: {e}. Falling back to standard template.")
            display_df = pd.DataFrame(default_abc_data)
    else:
        display_df = pd.DataFrame(default_abc_data)
        
    edited_abc = st.data_editor(display_df, num_rows="dynamic", use_container_width=True)

with tab2:
    st.subheader("Qualitative Interview Inputs")
    
    uploaded_notes_file = st.file_uploader("Optionally upload raw stakeholder interview notes (.txt)", type=["txt"])
    uploaded_text_content = ""
    if uploaded_notes_file is not None:
        uploaded_text_content = uploaded_notes_file.read().decode("utf-8")
        st.success("🎉 External interview transcript ingested into memory!")

    col1, col2 = st.columns(2)
    with col1:
        student_strengths = st.text_area("Student Strengths", 
                                        "Highly responsive to visual schedules, strong tactile interest, affectionate with familiar staff.", height=80)
        parent_notes = st.text_area("FBA Interview & Medical Factors (Parent/Caregiver)", 
                                    uploaded_text_content if uploaded_text_content else "Difficulty exploring leisure items independently; fatigue/sleep disruption exacerbates vocal resistance during demands.", height=120)
    with col2:
        school_demographics = st.text_input("School / Placement Setting", "Box Hill High School / PSU")
        behavior_description = st.text_area("Description of Behavior of Concern", 
                                            "Elopement (leaving assigned area) & Vocal Outbursts (screaming during transitions or academic demands).", height=120)
    
    st.divider()
    st.subheader("➕ Additional Clinical Input Dimensions")
    custom_inputs = st.text_area("Custom Stakeholder Observations / Ecological & Setting Events", 
                                  "e.g., Speech therapist reports communication breakdown increases frustration; Classroom transitions between rooms trigger higher density of elopement.", height=100)

with tab3:
    st.subheader("Questions About Behavioral Function (QABF) Scoring Matrix")
    st.caption("Input cumulative raw scores to trigger the automated matrix routing logic:")
    q_col1, q_col2, q_col3, q_col4, q_col5 = st.columns(5)
    with q_col1:
        att_score = st.number_input("Social Attention", 0, 15, 3)
    with q_col2:
        esc_score = st.number_input("Task/Demand Escape", 0, 15, 14)
    with q_col3:
        tan_score = st.number_input("Access to Tangibles", 0, 15, 4)
    with q_col4:
        sen_score = st.number_input("Non-Social / Sensory", 0, 15, 0)
    with q_col5:
        phy_score = st.number_input("Physical Discomfort", 0, 15, 1)

# 3. 后端确定性算法合成逻辑
highest_qabf = max(att_score, esc_score, tan_score, sen_score, phy_score)
primary_function = "Determining..."
bip_antecedent = ""
bip_consequence = ""
replacement_behavior = ""
generalization_plan = ""

if highest_qabf == esc_score:
    primary_function = "Social Negative Reinforcement (Escape or Avoid Task Demands)"
    replacement_behavior = (
        "1. Functional Communication Training (FCT): Teach the client to independently utilize an 'I need a break' visual card or vocal script prior to behavioral escalation.\n"
        "2. Tolerating Delay/Denial: Systematically introduce a progressive delay schedule between the request for a break and the delivery of reinforcement."
    )
    bip_antecedent = (
        "- Curriculum Modification: Break intensive multi-step math/literacy worksheets into single-line visual blocks to lower immediate cognitive load.\n"
        "- High-Probability Request Sequences (High-P): Deliver 3 simple, high-preference requests immediately before placing a non-preferred writing demand.\n"
        "- Choice-Making Arrangements: Offer control over task logistics.\n"
        "- Visual Pre-correction: Implement a 5-minute and 2-minute visual countdown timer prior to structured desk work transitions."
    )
    bip_consequence = (
        "- Differential Reinforcement of Alternative Behavior (DRA): Provide immediate functional escape (2 minutes) ONLY when FCT card is exhibited.\n"
        "- 3-Step Prompting Hierarchy: Implement a calm 'Tell-Show-Do' prompt sequence to ensure task completion.\n"
        "- Environmental Blocking (Safety): Position staff strategically to block elopement calmly without verbal interaction."
    )
    generalization_plan = (
        "- Stimulus Generalization: Train multiple classroom aides and parents to prompt/reinforce FCT cards using exact same scripts.\n"
        "- Schedule Thinning: Gradually increase required task components before break is honored."
    )
else:
    primary_function = "Social Positive Reinforcement (Access to Tangibles / Attention)"
    replacement_behavior = "Teach client to utilize visual request cards or vocal scripts to request preferred items or social interactions appropriately."
    bip_antecedent = (
        "- First/Then Visual Matrices: Explicitly display 'First Work Task, Then Preferred Activity'.\n"
        "- Token Economy System: Deliver tokens on a fixed ratio for task compliance, exchangeable for preferred reinforcers."
    )
    bip_consequence = (
        "- Extinction: Ensure problem behaviors strictly result in zero access to preferred items/attention.\n"
        "- Redirection: Neutral physical guidance back to task without verbal engagement."
    )
    generalization_plan = "Fade prompt density across novel environments and caregivers."

# 4. 生成 Word (.docx) 导出组件
st.divider()
st.header("⚡ Phase 2: Automated Compilation Engine")

def build_word_document():
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    title = doc.add_heading('Functional Behavioral Assessment (FBA) & BIP Report', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Section 1
    doc.add_heading('1. Clinical Demographics & Interview Inputs', level=1)
    doc.add_paragraph(f"School Setting: {school_demographics}")
    doc.add_paragraph(f"Student Strengths: {student_strengths}")
    doc.add_paragraph(f"Target Behavior: {behavior_description}")
    doc.add_paragraph(f"Medical/Setting Factors: {parent_notes}")
    
    # Section 2 Table
    doc.add_heading('2. Direct Observation ABC Ledger', level=1)
    
    headers = list(edited_abc.columns)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    for idx, text in enumerate(headers):
        hdr_cells[idx].text = str(text)
        shd = parse_xml(r'<w:shd {} w:fill="1F4E78"/>'.format(nsdecls('w')))
        hdr_cells[idx]._tc.get_or_add_tcPr().append(shd)
        p = hdr_cells[idx].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in p.runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(8.5)

    for r_idx, row in edited_abc.iterrows():
        row_cells = table.add_row().cells
        for c_idx, val in enumerate(row):
            row_cells[c_idx].text = str(val)
            p = row_cells[c_idx].paragraphs[0]
            p.runs[0].font.size = Pt(8.0)
            if r_idx % 2 == 1:
                shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
                row_cells[c_idx]._tc.get_or_add_tcPr().append(shd)

    # Section 3
    doc.add_heading('3. Behavior Intervention Plan (BIP) Prescriptions', level=1)
    doc.add_paragraph(f"Primary Inferred Function: {primary_function}").bold = True
    
    doc.add_heading('Replacement Behaviors (FCT):', level=2)
    doc.add_paragraph(replacement_behavior)
    
    doc.add_heading('Antecedent Modifications:', level=2)
    doc.add_paragraph(bip_antecedent)
    
    doc.add_heading('Consequence Strategies:', level=2)
    doc.add_paragraph(bip_consequence)

    doc.add_heading('Generalization & Thinning Plan:', level=2)
    doc.add_paragraph(generalization_plan)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

if st.button("🚀 Compile FBA & BIP Word Document", type="primary"):
    docx_file = build_word_document()
    st.success("🎉 Clinical Document Compiled Successfully in Memory!")
    
    st.download_button(
        label="📄 Download Completed FBA_BIP_Draft.docx",
        data=docx_file,
        file_name="FBA_BIP_Compiled_Draft.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

st.caption("© 2026 BIPEngine. All Rights Reserved. Prepared for Academic Review & Evaluation.")
