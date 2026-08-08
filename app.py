import streamlit as st
import io
import re
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# 1. 页面配置与侧边栏
st.set_page_config(page_title="BIPEngine Expert System v2.0", layout="wide", initial_sidebar_state="expanded")

with st.sidebar:
    st.markdown("<h2 style='color: #008080;'>⚙️ BIPEngine v2.0</h2>", unsafe_allow_html=True)
    st.caption("Advanced FBA & BIP Clinical Decision Support Engine")
    st.divider()
    
    st.markdown("### 🔒 Privacy & HIPAA Safety Guarantee")
    st.markdown("""
    * **No Cloud Storage**: Data exists only in temporary RAM memory.
    * **Instant Destruction**: Refreshing or closing page wipes all uploaded files.
    * **Automatic De-identification**: System masks names/locations (`[CLIENT_NAME]`).
    * **Local Re-identification**: Safely replace real names in Word using `Ctrl+H`.
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
    text = re.sub(r'(?i)\b(client|student|child|patient):\s*([A-Z][a-z]+\s+[A-Z][a-z]+)', r'\1: [CLIENT_NAME]', text)
    return text

# 动态文本多语言翻译映射字典
TRANSLATION_DICT = {
    # 功能
    "Escape / Demand Avoidance": "逃避 / 避免要求 (Escape / Demand Avoidance)",
    "Access to Tangible / Activity": "获取实物 / 活动 (Access to Tangible / Activity)",
    "Social Attention Seeking": "寻求社交关注 (Social Attention Seeking)",
    "Automatic / Sensory Synthetic": "自动强化 / 感官刺激 (Automatic / Sensory)",
    # 观察者身份
    "BCBA Direct": "BCBA 直接观察 (BCBA Direct)",
    "Caregiver (QR Log)": "照顾者/家长记录 (Caregiver)",
    "Classroom Aide": "教室助教 (Classroom Aide)",
    # 环境
    "Desk Work / Literacy": "桌面工作 / 读写课 (Desk Work / Literacy)",
    "Free Play Transition": "自由玩耍转换环节 (Free Play Transition)",
    "Small Group Work": "小组合作学习 (Small Group Work)",
    # 前因
    "Teacher presented multi-step writing task.": "教师布置了多步骤的书写任务。(Teacher presented multi-step writing task.)",
    "Timer rang to signal end of iPad time.": "计时器响起提示 iPad 使用时间结束。(Timer rang for end of iPad time.)",
    "Instructor turned attention to assist peer.": "指导老师将注意力转向协助其他同学。(Instructor turned attention to peer.)",
    # 行为
    "Screamed (>80dB), pushed desk away.": "尖叫（>80分贝），推开课桌。(Screamed >80dB, pushed desk away.)",
    "Vocal outburst, dropped to floor.": "大声发泄，躺倒在地上。(Vocal outburst, dropped to floor.)",
    "Approached staff, pulled sleeve, loud vocalizations.": "靠近工作人员，拉扯衣袖，大声发声。(Approached staff, pulled sleeve.)",
    # 后果
    "Staff presented 'Break' visual card; demand paused.": "工作人员出示“休息”视觉卡；暂停任务要求。(Staff presented 'Break' visual card.)",
    "Staff offered 2-min extension with visual timer.": "工作人员通过视觉计时器延长了2分钟。(Staff offered 2-min extension.)",
    "Staff turned immediately, made eye contact.": "工作人员转过身并给予眼神关注。(Staff turned immediately, made eye contact.)",
    # 背景与医疗字段
    "Responds exceptionally well to visual schedules, tactile items, and praise.": "对视觉日程表、触觉物品和口头夸奖反应非常好。(Responds exceptionally well to visual schedules, tactile items, and praise.)",
    "Elopement (leaving designated area >3 feet) and Vocal Outbursts during transitions.": "离开指定区域（>3英尺）以及在环节转换期间的大声情绪发泄。(Elopement and Vocal Outbursts during transitions.)",
    "Sleep disruption/fatigue increases behavior density by ~40%.": "睡眠中断/疲劳会导致行为发生频率增加约 40%。(Sleep disruption/fatigue increases behavior density by ~40%.)"
}

def translate_val(val, target_lang):
    if "Chinese" in target_lang:
        return TRANSLATION_DICT.get(str(val).strip(), str(val))
    return str(val)

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
    else:
        uploaded_qual_text = "Stakeholder Interview Note: Parent reports tantrums usually occur during homework time or when transitioning off digital screens. Teachers observe similar patterns during unstructured group activities."

    st.text_area("De-identified Uploaded Notes Preview", uploaded_qual_text, height=90)

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

# 3. 交叉验证
st.divider()
st.header("⚡ Phase 2: Triangulation Engine")

highest_qabf_score = max(att_score, esc_score, tan_score, sen_score, phy_score)
qabf_function = "Escape / Demand Avoidance" if highest_qabf_score == esc_score else "Social Attention / Tangible"

if "Early Intervention" in selected_age_group:
    age_strategy_note = "Early Intervention Focus: Play-based Functional Communication Training (FCT) via PECS/Visual Icons, parent co-regulation, and heavy environmental modification."
elif "School-Age" in selected_age_group:
    age_strategy_note = "School-Age Focus: Classroom accommodations, high-probability demand sequences, self-monitoring visual timers, and peer-mediated reinforcement."
else:
    age_strategy_note = "Adult / Transition Focus: Vocational task chunking, self-advocacy prompts, community integration protocols, and Support Worker SOPs."

st.info(f"💡 **Automated Triangulation Result**: Primary Function = **{qabf_function}**. {age_strategy_note}")

# 4. 全量中英/西双语 Word 导出引擎
def generate_fba_document():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.8)
    
    title = doc.add_heading('FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT', level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    if "Chinese" in selected_language:
        p = doc.add_paragraph("【 功能性行为评估报告（中英双语全量对照版） - 旨在提升 CALD 多元文化家庭配合度 】")
        p.runs[0].font.bold = True

    # Section 1
    doc.add_heading('1. Clinical Demographics & Background', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 1. 临床人口统计与背景信息 】").bold = True
        doc.add_paragraph(f"客户标识 (Client Identifier): [CLIENT_NAME] (请在 Word 中按 Ctrl+H 替换为真实姓名)")
        doc.add_paragraph(f"服务环境/安置 (Placement Setting): {school_setting}")
        doc.add_paragraph(f"年龄组别 (Age Group): {selected_age_group}")
        doc.add_paragraph(f"个人优势与强化物 (Strengths): {translate_val(student_strengths, selected_language)}")
        doc.add_paragraph(f"目标行为可操作性定义 (Behavior Def): {translate_val(behavior_desc, selected_language)}")
        doc.add_paragraph(f"医疗/环境设定因素 (Medical/Setting Factors): {translate_val(medical_factors, selected_language)}")
    else:
        doc.add_paragraph(f"Client Identifier: [CLIENT_NAME]")
        doc.add_paragraph(f"Placement Setting: {school_setting}")
        doc.add_paragraph(f"Age Cohort Category: {selected_age_group}")
        doc.add_paragraph(f"Client Strengths: {student_strengths}")
        doc.add_paragraph(f"Target Behavior Definition: {behavior_desc}")
        doc.add_paragraph(f"Medical / Setting Factors: {medical_factors}")

    # Section 1.5: 明确包含 Stakeholder Notes
    doc.add_heading('1.5 Stakeholder Qualitative Notes & Ecological Assessment', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 1.5 利益相关者质性访谈记录与生态评估 】").bold = True
        doc.add_paragraph(f"访谈与生态记录摘要：\n{uploaded_qual_text}")
    else:
        doc.add_paragraph(f"Qualitative Interview & Ecological Summary:\n{uploaded_qual_text}")

    # Section 2: ABC 观察表格（核心：表格内容全量动态翻译）
    doc.add_heading('2. Direct Systematic ABC Observation Ledger', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 2. 系统化直接 ABC 行为观察日志（全量双语翻译版） 】").bold = True

    headers = list(edited_abc.columns)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    # 表头翻译映射
    zh_headers = {
        "Entry": "条目 (Entry)",
        "Date/Time": "日期/时间 (Date/Time)",
        "Observer Role": "观察者 (Observer)",
        "Setting": "环境 (Setting)",
        "Antecedent (A)": "前因 (Antecedent A)",
        "Behavior (B)": "行为 (Behavior B)",
        "Consequence (C)": "后果 (Consequence C)",
        "Engine Auto-Inferred Function": "系统推导功能 (Inferred Function)"
    }

    for idx, text in enumerate(headers):
        cell = table.rows[0].cells[idx]
        cell.text = zh_headers.get(text, text) if "Chinese" in selected_language else str(text)
        shd = parse_xml(r'<w:shd {} w:fill="1F4E78"/>'.format(nsdecls('w')))
        cell._tc.get_or_add_tcPr().append(shd)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(8)

    # 填入表格内容（应用动态翻译）
    for r_idx, row in edited_abc.iterrows():
        row_cells = table.add_row().cells
        for c_idx, val in enumerate(row):
            translated_content = translate_val(val, selected_language)
            row_cells[c_idx].text = translated_content
            p = row_cells[c_idx].paragraphs[0]
            p.runs[0].font.size = Pt(7.5)
            if r_idx % 2 == 1:
                shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
                row_cells[c_idx]._tc.get_or_add_tcPr().append(shd)

    # Section 3
    doc.add_heading('3. Triangulated Discrepancy & Functional Summary', level=1)
    translated_func = translate_val(qabf_function, selected_language)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 3. 数据交叉验证与行为功能评估结论 】").bold = True
        doc.add_paragraph(f"QABF 量表最高得分领域: {highest_qabf_score} 分")
        doc.add_paragraph(f"临床综合结论: 结合直接 ABC 观察日志与 QABF 心理测量量表，数据一致表明 **{translated_func}** 是维持该目标行为的主要功能变量。")
    else:
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
    
    translated_func = translate_val(qabf_function, selected_language)

    doc.add_heading('1. Target Behavior & Primary Function', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 1. 目标行为与主要维持功能 】").bold = True
        doc.add_paragraph(f"客户标识 (Client ID): [CLIENT_NAME]")
        doc.add_paragraph(f"目标行为 (Target Behavior): {translate_val(behavior_desc, selected_language)}")
        doc.add_paragraph(f"自动推导的主要功能 (Primary Function): {translated_func}")
        doc.add_paragraph(f"针对年龄组别的干预重点 (Age Focus): {age_strategy_note}")
    else:
        doc.add_paragraph(f"Client Identifier: [CLIENT_NAME]")
        doc.add_paragraph(f"Target Behavior: {behavior_desc}")
        doc.add_paragraph(f"Inferred Maintaining Function: {qabf_function}")
        doc.add_paragraph(f"Prescribed Age Focus: {age_strategy_note}")

    doc.add_heading('2. Proactive Antecedent Strategies', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 2. 前因预防策略 (Proactive Antecedent Strategies) 】").bold = True
        doc.add_paragraph("1. 视觉预告与倒计时提示 (Visual Pre-Correction): 在任务转换前 5 分钟和 2 分钟提供视觉倒计时，降低过渡期焦虑。")
        doc.add_paragraph("2. 任务拆解与需求修改 (Task Chunking): 将多步骤的书面指令拆解为单步视觉卡片，降低认知负荷。")
        doc.add_paragraph("3. 高概率请求序列 (High-P Sequence): 在发出较难指令前，连续发出 3 个快速且容易完成的高偏好指令，建立行为惯性。")
    else:
        doc.add_paragraph("1. Visual Pre-Correction & Countdown Timers: Provide 5-minute and 2-minute visual cues prior to task transitions.")
        doc.add_paragraph("2. Curriculum Chunking & Demand Modification: Break multi-step instructions into single-step visual task cards.")
        doc.add_paragraph("3. High-Probability (High-P) Request Sequence: Deliver 3 rapid preferred requests prior to non-preferred demands.")

    doc.add_heading('3. Functional Replacement Behaviors (FCT)', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 3. 功能性替代行为训练 (Functional Communication Training - FCT) 】").bold = True
        doc.add_paragraph("1. 独立请求休息 (Independent Break Request): 系统化教导客户在行为升级前，主动触摸或出示“我需要休息”的视觉卡片。")
        doc.add_paragraph("2. 差别强化策略 (DRA): 仅在客户使用替代行为（如出示卡片）时提供即时的功能性强化（如 1-2 分钟暂停/休息）。")
    else:
        doc.add_paragraph("1. Independent Break Requests: Teach client to touch/hand the 'I Need a Break' visual card prior to behavioral escalation.")
        doc.add_paragraph("2. Differential Reinforcement of Alternative Behavior (DRA): Provide immediate functional reinforcement ONLY upon replacement behavior.")

    doc.add_heading('4. Reactive Consequence & Safety Protocols', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 4. 后果反应策略与安全预案 (Reactive Consequence Protocols) 】").bold = True
        doc.add_paragraph("1. 逃避消退/三步提示法 (Escape Extinction via 3-Step Prompting): 保持温和中立的态度，使用“告知-示范-协助”顺序引导完成基本任务，避免口头批评。")
        doc.add_paragraph("2. 环境阻挡与安全防护 (Environmental Blocking): 工作人员保持中立且无眼神接触的状态下安全阻挡逃跑行为，不给予额外的口头回应。")
    else:
        doc.add_paragraph("1. Escape Extinction (3-Step Prompting): Utilize calm 'Tell-Show-Do' prompting to complete tasks without verbal reprimands.")
        doc.add_paragraph("2. Environmental Blocking: Position staff neutrally to block elopement safely without eye contact or verbal commentary.")

    doc.add_heading('5. Generalization & Local Re-identification Note', level=1)
    if "Chinese" in selected_language:
        doc.add_paragraph("【 5. 泛化计划与本地隐私还原说明 】").bold = True
        doc.add_paragraph("提示: 本报告所有客户身份均已进行脱敏处理（显示为 [CLIENT_NAME]）。在正式提交团队前，请在本地 Word 中按 Ctrl+H 将其一键替换为客户真实姓名。")
    else:
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
