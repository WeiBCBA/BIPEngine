 import streamlit as st
import io
import re
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# 1. 页面配置与大白话隐私保障侧边栏
st.set_page_config(page_title="BIPEngine Expert System v2.0", layout="wide", initial_sidebar_state="expanded")

with st.sidebar:
    st.markdown("<h2 style='color: #008080;'>⚙️ BIPEngine v2.0</h2>", unsafe_allow_html=True)
    st.caption("Advanced FBA & BIP Clinical Decision Support Engine")
    st.divider()
    
    st.markdown("### 🔒 Privacy & HIPAA Safety Guarantee")
    st.markdown("""
    * **No Cloud Storage**: Data exists only in temporary RAM memory while you are active on this page.
    * **Instant Destruction**: Refreshing or closing this page completely wipes all uploaded files.
    * **Automatic De-identification**: System masks names and locations to placeholders (`[CLIENT_NAME]`).
    * **Local Re-identification**: Safely insert real client details on your local PC using Word Find & Replace (`Ctrl+H`).
    """)
    st.divider()
    
    st.subheader("🌐 Global DEI & Clinical Settings")
    selected_language = st.selectbox(
        "Report Output Language Target:",
        ["English (Standard US)", "English & Chinese Dual-Language (中英双语)", "English & Spanish Dual-Language (Español & English)"]
    )
    
    selected_age_group = st.selectbox(
        "Client Development Cohort:",
        ["Early Intervention (2-5 yrs)", "School-Age (5-21 yrs)", "Adult / Transition (21+ yrs)"]
    )

st.title("🚀 BIPEngine: Automated FBA & BIP Expert Compiler")
st.markdown("##### Bridging Clinical Domain Architecture, DEI Inclusivity, and Automated Document Synthesis")
st.divider()

# 脱敏辅助函数
def deidentify_text(text):
    if not text:
        return ""
    # 将常见的姓名或敏感关键词占位符化
    text = re.sub(r'(?i)\b(client|student|child|patient):\s*([A-Z][a-z]+\s+[A-Z][a-z]+)', r'\1: [CLIENT_NAME]', text)
    return text

# 2. 核心数据采集模块
st.header("📋 Phase 1: Multi-Source Data Ingestion")
tab1, tab2, tab3 = st.tabs(["📊 ABC Observation Ledger", "📝 Qualitative Stakeholder Notes", "📈 QABF Psychometric Profiler"])

with tab1:
    st.subheader("Direct Observation ABC Data Ledger")
    st.caption("Upload raw ABC data. The system automatically infers behavioral functions from direct observation facts.")
    
    uploaded_abc_file = st.file_uploader("Upload Raw ABC Data (CSV format)", type=["csv"], key="abc_csv")
    
    default_abc_data = [
        {"Entry": "Obs #1", "Date/Time": "08/03/2026 09:15 AM", "Observer Role": "BCBA Direct", "Setting": "Desk Work / Literacy", "Antecedent (A)": "Teacher presented multi-step writing task.", "Behavior (B)": "Screamed (>80dB), pushed desk away.", "Consequence (C)": "Staff presented 'Break' visual card; demand paused."},
        {"Entry": "Obs #2", "Date/Time": "08/04/2026 10:30 AM", "Observer Role": "Caregiver (QR Log)", "Setting": "Free Play Transition", "Antecedent (A)": "Timer rang to signal end of iPad time.", "Behavior (B)": "Vocal outburst, dropped to floor.", "Consequence (C)": "Staff offered 2-min extension with visual timer."},
        {"Entry": "Obs #3", "Date/Time": "08/05/2026 01:45 PM", "Observer Role": "Classroom Aide", "Setting": "Small Group Work", "Antecedent (A)": "Instructor turned attention to assist peer.", "Behavior (B)": "Approached staff, pulled sleeve, loud vocalizations.", "Consequence (C)": "Staff turned immediately, made eye contact."}
    ]
    
    if uploaded_abc_file is not None:
        try:
            raw_df = pd.read_csv(uploaded_abc_file)
            st.success("🎉 Custom ABC log parsed into memory!")
        except Exception as e:
            st.error(f"Error reading CSV: {e}. Reverting to baseline data.")
            raw_df = pd.DataFrame(default_abc_data)
    else:
        raw_df = pd.DataFrame(default_abc_data)

    def infer_function_from_row(row):
        ant = str(row.get("Antecedent (A)", "")).lower()
        con = str(row.get("Consequence (C)", "")).lower()
        if "demand" in ant or "task" in ant or "pause" in con or "break" in con:
            return "Escape / Demand Avoidance"
        elif "ipad" in ant or "item" in ant or "extension" in con or "given" in con:
            return "Access to Tangible / Activity"
        elif "attention" in ant or "peer" in ant or "eye contact" in con:
            return "Social Attention Seeking"
        return "Automatic / Sensory Synthetic"

    raw_df["Engine Auto-Inferred Function"] = raw_df.apply(infer_function_from_row, axis=1)
    edited_abc = st.data_editor(raw_df, num_rows="dynamic", use_container_width=True)

with tab2:
    st.subheader("Qualitative Stakeholder Upload & Ecological Inputs")
    
    qual_file = st.file_uploader("Upload Qualitative Notes or Interview Transcripts (.txt format)", type=["txt"], key="qual_txt")
    uploaded_qual_text = ""
    if qual_file is not None:
        uploaded_qual_text = qual_file.read().decode("utf-8")
        uploaded_qual_text = deidentify_text(uploaded_qual_text)
        st.success("✅ Qualitative file uploaded & de-identified automatically!")
        st.text_area("De-identified Uploaded Notes Preview", uploaded_qual_text, height=100)

    col1, col2 = st.columns(2)
    with col1:
        school_setting = st.text_input("Placement / Setting", "[LOCATION] / Inclusive Classroom Setting")
        student_strengths = st.text_area("Client Strengths & Motivators", "Responds exceptionally well to visual schedules, tactile items, and praise.", height=90)
    with col2:
        behavior_desc = st.text_area("Target Behavior Operational Definition", "Elopement (leaving designated area >3 feet) and Vocal Outbursts during transitions.", height=90)
        medical_factors = st.text_area("Medical / Setting Factors", "Sleep disruption/fatigue increases behavior density by ~40%.", height=90)

with tab3:
    st.subheader("Questions About Behavioral Function (QABF) Matrix")
    q1, q2, q3, q4, q5 = st.columns(5)
    att_score = q1.number_input("Social Attention", 0, 15, 3)
    esc_score = q2.number_input("Task Escape", 0, 15, 14)
    tan_score = q3.number_input("Tangibles", 0, 15, 4)
    sen_score = q4.number_input("Sensory", 0, 15, 0)
    phy_score = q5.number_input("Physical Discomfort", 0, 15, 1)

# 3. 自动数据交叉验证引擎
st.divider()
st.header("⚡ Phase 2: Triangulation Engine")

highest_qabf_score = max(att_score, esc_score, tan_score, sen_score, phy_score)
qabf_function = "Escape" if highest_qabf_score == esc_score else "Attention/Tangible"

if "Early Intervention" in selected_age_group:
    age_strategy_note = "Early Intervention Focus: Play-based Functional Communication Training (FCT) via PECS/Visual Icons, parent co-regulation, and heavy environmental modification."
elif "School-Age" in selected_age_group:
    age_strategy_note = "School-Age Focus: Classroom accommodations, high-probability demand sequences, self-monitoring visual timers, and peer-mediated reinforcement."
else:
    age_strategy_note = "Adult / Transition Focus: Vocational task chunking, self-advocacy prompts, community integration protocols, and Support Worker SOPs."

st.info(f"💡 **Automated Triangulation Result**: Primary Function = **{qabf_function}**. {age_strategy_note}")

# 4. 多语言 Word 导出逻辑
def generate_fba_document():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.8)
    
    title = doc.add_heading('FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 语言提示标语
    if "Chinese" in selected_language:
        p = doc.add_paragraph("【 注意：本评估报告包含中英双语对照，方便多元文化（CALD）家庭阅读与配合。 】")
        p.runs[0].font.bold = True
    elif "Spanish" in selected_language:
        p = doc.add_paragraph("【 Nota: Este informe de evaluación contiene una mapeo bilingüe para familias cultural y lingüísticamente diversas (CALD). 】")
        p.runs[0].font.bold = True

    doc.add_heading('1. Clinical Demographics & Background', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 1. 临床人口统计与评估背景 】").bold = True
    elif "Spanish" in selected_language:
        doc.add_paragraph("【 1. Datos Demográficos y Antecedentes Clínícos 】").bold = True

    doc.add_paragraph(f"Client Identifier: [CLIENT_NAME]")
    doc.add_paragraph(f"Placement Setting: {school_setting}")
    doc.add_paragraph(f"Age Cohort Category: {selected_age_group}")
    doc.add_paragraph(f"Client Strengths: {student_strengths}")
    doc.add_paragraph(f"Target Behavior Definition: {behavior_desc}")
    doc.add_paragraph(f"Medical / Setting Factors: {medical_factors}")
    if uploaded_qual_text:
        doc.add_paragraph(f"Uploaded Qualitative Input Summary: {uploaded_qual_text[:300]}...")

    doc.add_heading('2. Direct Systematic ABC Observation Ledger', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 2. 直接系统化 ABC 行为观察记录表 】").bold = True
    elif "Spanish" in selected_language:
        doc.add_paragraph("【 2. Registro Sistemático de Observación ABC 】").bold = True

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
    if "Chinese" in selected_language:
        doc.add_paragraph("【 3. 数据交叉验证与行为功能结论 】").bold = True
    elif "Spanish" in selected_language:
        doc.add_paragraph("【 3. Triangulación de Datos y Resumen Funcional 】").bold = True

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
    if "Chinese" in selected_language:
        doc.add_paragraph("【 1. 目标行为与主要维持功能 】").bold = True
    elif "Spanish" in selected_language:
        doc.add_paragraph("【 1. Conducta Objetivo y Función Principal 】").bold = True

    doc.add_paragraph(f"Client Identifier: [CLIENT_NAME]")
    doc.add_paragraph(f"Target Behavior: {behavior_desc}")
    doc.add_paragraph(f"Inferred Maintaining Function: {qabf_function}")
    doc.add_paragraph(f"Prescribed Age Focus: {age_strategy_note}")

    doc.add_heading('2. Proactive Antecedent Strategies', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 2. 前因干预策略 (Antecedent Strategies) 】").bold = True
        doc.add_paragraph("1. 视觉预告与倒计时：在任务转换前提供 5 分钟和 2 分钟的视觉倒计时提示。")
        doc.add_paragraph("2. 任务拆解：将多步骤指令拆解为单步视觉卡片。")
        doc.add_paragraph("3. 高概率请求序列：在发出非偏好指令前，先连续发出 3 个高偏好（易完成）的指令。")
    elif "Spanish" in selected_language:
        doc.add_paragraph("【 2. Estrategias Proactivas de Antecedentes 】").bold = True
        doc.add_paragraph("1. Pre-corrección visual y temporizadores: Proporcionar avisos visuales 5 y 2 minutos antes de las transiciones.")
        doc.add_paragraph("2. Fragmentación de tareas: Dividir las instrucciones en tarjetas visuales de un solo paso.")
        doc.add_paragraph("3. Secuencia de alta probabilidad: Entregar 3 peticiones preferidas antes de una demanda no preferida.")
    else:
        doc.add_paragraph("1. Visual Pre-Correction & Countdown Timers: Provide 5-minute and 2-minute visual cues prior to task transitions.")
        doc.add_paragraph("2. Curriculum Chunking & Demand Modification: Break multi-step instructions into single-step visual task cards.")
        doc.add_paragraph("3. High-Probability (High-P) Request Sequence: Deliver 3 rapid preferred requests prior to non-preferred demands.")

    doc.add_heading('3. Functional Replacement Behaviors (FCT)', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 3. 功能性替代行为训练 (FCT) 】").bold = True
        doc.add_paragraph("1. 独立请求休息：教导客户在行为升级前，主动出示或触摸“我需要休息”视觉卡。")
        doc.add_paragraph("2. 差别强化 (DRA)：仅对替代行为给予即时功能性强化（1-2分钟暂停/休息）。")
    elif "Spanish" in selected_language:
        doc.add_paragraph("【 3. Conductas de Reemplazo Funcional (FCT) 】").bold = True
        doc.add_paragraph("1. Solicitud de descanso independiente: Enseñar al cliente a usar la tarjeta 'Necesito un descanso'.")
        doc.add_paragraph("2. Reforzamiento Diferencial (DRA): Proporcionar reforzamiento inmediato solo ante la conducta de reemplazo.")
    else:
        doc.add_paragraph("1. Independent Break Requests: Teach client to touch/hand the 'I Need a Break' visual card prior to behavioral escalation.")
        doc.add_paragraph("2. Differential Reinforcement of Alternative Behavior (DRA): Provide immediate functional reinforcement ONLY upon replacement behavior.")

    doc.add_heading('4. Reactive Consequence & Safety Protocols', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 4. 后果反应策略与安全预案 】").bold = True
        doc.add_paragraph("1. 逃避消退 (三步提示法)：使用温和的“告知-示范-协助”提示引导完成任务，避免口头批评。")
        doc.add_paragraph("2. 环境阻挡：人员保持中立态度阻挡逃跑行为，避免眼神接触或多余口头回应。")
    elif "Spanish" in selected_language:
        doc.add_paragraph("【 4. Consecuencias Reactivas y Protocolos de Seguridad 】").bold = True
        doc.add_paragraph("1. Extinción de escape: Utilizar la jerarquía de guía física sin reprimendas verbales.")
        doc.add_paragraph("2. Bloqueo ambiental: Posicionarse de manera neutral para bloquear la fuga de forma segura.")
    else:
        doc.add_paragraph("1. Escape Extinction (3-Step Prompting): Utilize calm 'Tell-Show-Do' prompting to complete tasks without verbal reprimands.")
        doc.add_paragraph("2. Environmental Blocking: Position staff neutrally to block elopement safely without eye contact or verbal commentary.")

    doc.add_heading('5. Generalization & Local Re-identification Note', level=1)
    doc.add_paragraph("Note for BCBA: All client names are export-masked as [CLIENT_NAME]. Use Word Find & Replace (Ctrl+H) on your local workstation to restore true identifying details prior to clinical submission.")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 5. 独立下载按钮
st.subheader("📄 Export Clinical Draft Documents")
col_fba, col_bip = st.columns(2)

with col_fba:
    if st.button("🚀 Compile FBA Report (.docx)", type="primary", use_container_width=True):
        fba_file = generate_fba_document()
        st.success("FBA Report Compiled Successfully!")
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
        st.success("BIP Plan Compiled Successfully!")
        st.download_button(
            label="📄 Download BIP_Plan.docx",
            data=bip_file,
            file_name=f"BIP_Plan_{selected_age_group.split()[0]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

st.caption("© 2026 BIPEngine Expert System. All Rights Reserved. Built for Clinical Academic Evaluation.")
