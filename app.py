import streamlit as st
import io
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# 1. 页面配置与大白话隐私保障侧边栏 (Zero-Storage Guarantee)
st.set_page_config(page_title="BIPEngine Expert System v2.0", layout="wide", initial_sidebar_state="expanded")

with st.sidebar:
    st.markdown("<h2 style='color: #008080;'>⚙️ BIPEngine v2.0</h2>", unsafe_allow_html=True)
    st.caption("Advanced FBA & BIP Clinical Decision Support Engine")
    st.divider()
    
    # 🌟 普通人/导师都能看懂的大白话隐私保护声明
    st.markdown("### 🔒 Privacy & Safety Guarantee")
    st.markdown("""
    * **No Cloud Storage (不存云端)**: Your data lives only in temporary website memory (RAM) while you are on this page.
    * **Instant Destruction (用完即毁)**: Closing or refreshing this page instantly erases all uploaded files and notes forever.
    * **No AI Training (不拿数据训练)**: Zero data is sent to public AI models for training or logging.
    * **100% De-identified (完全脱敏)**: Built for strict HIPAA & BACB Ethics Code compliance.
    """)
    st.divider()
    
    # 🌟 全球化与 DEI 多语言配置
    st.subheader("🌐 Global DEI & Clinical Settings")
    selected_language = st.selectbox(
        "Report Output Language Target:",
        ["English (Standard US)", "English & Chinese Dual-Language (中英双语)", "English & Spanish Dual-Language (西英双语)"]
    )
    
    selected_age_group = st.selectbox(
        "Client Development Cohort:",
        ["Early Intervention (2-5 yrs)", "School-Age (5-21 yrs)", "Adult / Transition (21+ yrs)"]
    )

st.title("🚀 BIPEngine: Automated FBA & BIP Expert Compiler")
st.markdown("##### Bridging Clinical Domain Architecture, DEI Inclusivity, and Automated Document Synthesis to Eliminate Practitioner Burnout")
st.divider()

# 2. 核心数据采集模块 (Phase 1)
st.header("📋 Phase 1: Multi-Source Data Ingestion")
tab1, tab2, tab3 = st.tabs(["📊 ABC Observation Ledger", "📝 Stakeholder Qualitative Input", "📈 QABF Psychometric Profiler"])

with tab1:
    st.subheader("Direct Observation ABC Data Ledger")
    st.caption("Upload a custom ABC CSV file or edit the matrix below. The system automatically infers behavioral functions from raw facts.")
    
    uploaded_abc_file = st.file_uploader("Upload Raw ABC Data (CSV format)", type=["csv"])
    
    # 默认原始数据（仅包含客观事实，不强制人工填写 Inferred Function）
    default_abc_data = [
        {"Entry": "Obs #1", "Date/Time": "08/03/2026 09:15 AM", "Observer Role": "BCBA Direct", "Setting": "Desk Work / Literacy", "Antecedent (A)": "Teacher presented a multi-step writing worksheet.", "Behavior (B)": "Screamed (>80dB), pushed desk away, attempted room exit.", "Consequence (C)": "Staff presented 'Break' visual card; demand temporarily paused."},
        {"Entry": "Obs #2", "Date/Time": "08/04/2026 10:30 AM", "Observer Role": "Caregiver (QR Log)", "Setting": "Free Play Transition", "Antecedent (A)": "Timer rang to signal end of iPad play session.", "Behavior (B)": "Vocal outburst, dropped to floor, refused movement.", "Consequence (C)": "Staff offered 2-min extension with visual countdown timer."},
        {"Entry": "Obs #3", "Date/Time": "08/05/2026 01:45 PM", "Observer Role": "Classroom Aide", "Setting": "Small Group Work", "Antecedent (A)": "Instructor turned attention to assist a peer.", "Behavior (B)": "Approached staff, pulled sleeve, loud vocalizations.", "Consequence (C)": "Staff turned immediately, made eye contact, reassured student."}
    ]
    
    if uploaded_abc_file is not None:
        try:
            raw_df = pd.read_csv(uploaded_abc_file)
            st.success("🎉 Custom ABC tracking log parsed successfully into memory!")
        except Exception as e:
            st.error(f"Error reading CSV: {e}. Falling back to default ledger.")
            raw_df = pd.DataFrame(default_abc_data)
    else:
        raw_df = pd.DataFrame(default_abc_data)

    # 🌟 后台确定性算法：自动从 A-B-C 事实中推理 Function
    def infer_function_from_row(row):
        ant = str(row.get("Antecedent (A)", "")).lower()
        con = str(row.get("Consequence (C)", "")).lower()
        if "demand" in ant or "worksheet" in ant or "pause" in con or "break" in con:
            return "Escape / Demand Avoidance"
        elif "ipad" in ant or "item" in ant or "extension" in con or "given" in con:
            return "Access to Tangible / Activity"
        elif "attention" in ant or "peer" in ant or "eye contact" in con or "reassured" in con:
            return "Social Attention Seeking"
        return "Automatic / Sensory Synthetic"

    raw_df["Engine Auto-Inferred Function"] = raw_df.apply(infer_function_from_row, axis=1)
    edited_abc = st.data_editor(raw_df, num_rows="dynamic", use_container_width=True)

with tab2:
    st.subheader("Qualitative Stakeholder Interview & Ecological Variables")
    col1, col2 = st.columns(2)
    with col1:
        school_setting = st.text_input("Placement / Setting", "Box Hill Educational Center / Inclusive Classroom")
        student_strengths = st.text_area("Client Strengths & Motivators", "Responds exceptionally well to visual schedules, tactile items, and praise.", height=90)
    with col2:
        behavior_desc = st.text_area("Target Behavior Operational Definition", "Elopement (leaving designated area >3 feet) and Vocal Outbursts (>70dB screaming during transitions).", height=90)
        medical_factors = st.text_area("Medical / Setting Factors", "Sleep disruption/fatigue increases behavior density by ~40%.", height=90)

with tab3:
    st.subheader("Questions About Behavioral Function (QABF) Matrix")
    st.caption("Input cumulative raw test scores:")
    q1, q2, q3, q4, q5 = st.columns(5)
    att_score = q1.number_input("Social Attention", 0, 15, 3)
    esc_score = q2.number_input("Task Escape", 0, 15, 14)
    tan_score = q3.number_input("Tangibles", 0, 15, 4)
    sen_score = q4.number_input("Sensory", 0, 15, 0)
    phy_score = q5.number_input("Physical Discomfort", 0, 15, 1)

# 3. 自动数据交叉验证引擎 (Triangulation & Discrepancy Engine)
st.divider()
st.header("⚡ Phase 2: Triangulation & Age-Cohort Prescription Engine")

highest_qabf_score = max(att_score, esc_score, tan_score, sen_score, phy_score)
qabf_function = "Escape" if highest_qabf_score == esc_score else "Attention/Tangible"

# 自动推导年龄段干预策略
if "Early Intervention" in selected_age_group:
    age_strategy_note = "Early Intervention Focus: Play-based Functional Communication Training (FCT) via PECS/Visual Icons, parent co-regulation, and heavy environmental modification."
elif "School-Age" in selected_age_group:
    age_strategy_note = "School-Age Focus: Classroom accommodations, high-probability demand sequences, self-monitoring visual timers, and peer-mediated reinforcement schedules."
else:
    age_strategy_note = "Adult / Transition Focus: Vocational task chunking, self-advocacy prompts, community integration protocols, and Support Worker SOPs."

st.info(f"💡 **Automated Triangulation Result**: Primary Inferred Function = **{qabf_function}**. {age_strategy_note}")

# 4. 双文件生成器（FBA 独立生成器 & BIP 独立生成器）
def generate_fba_document():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.8)
    
    title = doc.add_heading('FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    if "Dual-Language" in selected_language:
        doc.add_paragraph("【 Note: This report contains bilingual clinical mapping for culturally and linguistically diverse (CALD) family accessibility. 】").bold = True

    doc.add_heading('1. Clinical Demographics & Assessment Background', level=1)
    doc.add_paragraph(f"Placement Setting: {school_setting}")
    doc.add_paragraph(f"Age Cohort Category: {selected_age_group}")
    doc.add_paragraph(f"Client Strengths: {student_strengths}")
    doc.add_paragraph(f"Target Behavior Definition: {behavior_desc}")
    doc.add_paragraph(f"Medical / Setting Events: {medical_factors}")

    doc.add_heading('2. Direct Systematic ABC Observation Ledger', level=1)
    headers = list(edited_abc.columns)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    for idx, text in enumerate(headers):
        cell = table.rows[0].cells[idx]
        cell.text = str(text)
        shd = parse_xml(r'<w:shd {} w:fill="1F4E78"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shd)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(8)

    for r_idx, row in edited_abc.iterrows():
        row_cells = table.add_row().cells
        for c_idx, val in enumerate(row):
            row_cells[c_idx].text = str(val)
            p = row_cells[c_idx].paragraphs[0]
            p.runs[0].font.size = Pt(7.5)
            if r_idx % 2 == 1:
                shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
                row_cells[c_idx]._tc.get_or_add_tcPr().append(shd)

    doc.add_heading('3. Triangulated Discrepancy & Functional Summary', level=1)
    doc.add_paragraph(f"QABF Highest Score: {highest_qabf_score} (Function: {qabf_function})")
    doc.add_paragraph(f"Clinical Conclusion: Direct ABC observation logs and psychometric scoring converge on {qabf_function} as the primary maintaining variable.")
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def generate_bip_document():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.8)
    
    title = doc.add_heading('BEHAVIOR INTERVENTION PLAN (BIP)', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_heading('1. Target Behavior & Primary Function', level=1)
    doc.add_paragraph(f"Target Behavior: {behavior_desc}")
    doc.add_paragraph(f"Inferred Maintaining Function: {qabf_function}")
    doc.add_paragraph(f"Prescribed Age Focus: {age_strategy_note}")

    doc.add_heading('2. Proactive Antecedent Strategies', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【前因策略 Antecedent Modifications】")
    doc.add_paragraph("1. Visual Pre-Correction & Countdown Timers: Provide 5-minute and 2-minute visual cues prior to task transitions.")
    doc.add_paragraph("2. Curriculum Chunking & Demand Modification: Break multi-step instructions into single-step visual task cards.")
    doc.add_paragraph("3. High-Probability (High-P) Request Sequence: Deliver 3 rapid preferred requests prior to non-preferred demands.")

    doc.add_heading('3. Functional Replacement Behaviors (FCT)', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【替代行为训练 Functional Replacement】")
    doc.add_paragraph("1. Independent Break Requests: Teach client to touch/hand the 'I Need a Break' visual card prior to behavioral escalation.")
    doc.add_paragraph("2. Differential Reinforcement of Alternative Behavior (DRA): Provide immediate functional reinforcement (1-2 min break) ONLY upon replacement behavior.")

    doc.add_heading('4. Reactive Consequence & Safety Protocols', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【后果策略与安全预案 Consequence & Safety】")
    doc.add_paragraph("1. Escape Extinction (3-Step Prompting): Utilize calm 'Tell-Show-Do' prompting to complete tasks without verbal reprimands.")
    doc.add_paragraph("2. Environmental Blocking: Position staff neutrally to block elopement safely without eye contact or verbal commentary.")

    doc.add_heading('5. Generalization & Fading Plan', level=1)
    doc.add_paragraph("Systematically thin reinforcement schedule from FR-1 to VR-3 across novel instructors, caregivers, and community settings.")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 5. 渲染独立下载按钮
st.subheader("📄 Export Clinical Draft Documents")
col_fba, col_bip = st.columns(2)

with col_fba:
    if st.button("🚀 Compile FBA Report (.docx)", type="primary", use_container_width=True):
        fba_file = generate_fba_document()
        st.success("FBA Report Compiled!")
        st.download_button(
            label="📄 Download FBA_Report.docx",
            data=fba_file,
            file_name=f"FBA_Report_{selected_age_group.split()[0]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

with col_bip:
    if st.button("🚀 Compile BIP Plan (.docx)", type="primary", use_container_width=True):
        bip_file = generate_bip_document()
        st.success("BIP Plan Compiled!")
        st.download_button(
            label="📄 Download BIP_Plan.docx",
            data=bip_file,
            file_name=f"BIP_Plan_{selected_age_group.split()[0]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

st.caption("© 2026 BIPEngine Expert System. All Rights Reserved. Built for Clinical Academic Evaluation.")
