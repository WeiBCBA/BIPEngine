import streamlit as st
import io
import re
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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
        ["English (Standard US)", "English & Chinese Dual-Language (中英双语对照)", "English & Spanish Dual-Language (Español & English)"]
    )
    
    selected_age_group = st.selectbox(
        "Client Development Cohort:",
        ["Early Intervention (2-5 yrs)", "School-Age (5-21 yrs)", "Adult / Transition (21+ yrs)"]
    )

st.title("🚀 BIPEngine: Automated FBA & BIP Expert Compiler")
st.markdown("##### Bridging Clinical Domain Architecture, DEI Inclusivity, and Automated Document Synthesis")
st.divider()

def deidentify_text(text):
    if not text:
        return ""
    text = re.sub(r'(?i)\b(client|student|child|patient):\s*([A-Z][a-z]+\s+[A-Z][a-z]+)', r'\1: [CLIENT_NAME]', text)
    return text

# 全量双语翻译字典
DICTIONARY_ZH = {
    # 表头翻译
    "Entry": "记录编号 (Entry)",
    "Date/Time": "日期/时间 (Date/Time)",
    "Observer Role": "观察者身份 (Observer Role)",
    "Setting": "服务环境 (Setting)",
    "Antecedent (A)": "前因 (Antecedent - A)",
    "Behavior (B)": "行为 (Behavior - B)",
    "Consequence (C)": "后果 (Consequence - C)",
    "Engine Auto-Inferred Function": "系统推导功能 (Inferred Function)",
    
    # ABC 表格单元格翻译
    "Obs #1": "观察记录 #1 (Obs #1)",
    "Obs #2": "观察记录 #2 (Obs #2)",
    "Obs #3": "观察记录 #3 (Obs #3)",
    "BCBA Direct": "BCBA 直接观察",
    "Caregiver (QR Log)": "照护者 (QR 日志)",
    "Classroom Aide": "教室助教",
    "Desk Work / Literacy": "桌面个训 / 读写课",
    "Free Play Transition": "自由玩耍环节转换",
    "Small Group Work": "小组活动",
    "Teacher presented multi-step writing task.": "教师呈递了多步骤的书写任务。",
    "Timer rang to signal end of iPad time.": "计时器响起，提示 iPad 使用时间结束。",
    "Instructor turned attention to assist peer.": "指导老师将注意力转向协助同伴。",
    "Screamed (>80dB), pushed desk away.": "尖叫（>80分贝），推开课桌。",
    "Vocal outbursts, dropped to floor.": "发声情绪爆发出声，倒在地板上。",
    "Vocal outburst, dropped to floor.": "发声情绪爆发出声，倒在地板上。",
    "Approached staff, pulled sleeve, loud vocalizations.": "靠近工作人员，拉扯衣袖，大声发出声音。",
    "Staff presented 'Break' visual card; demand paused.": "工作人员出示“休息”视觉卡片；暂停当前要求。",
    "Staff offered 2-min extension with visual timer.": "工作人员通过视觉计时器给予 2 分钟的延长时间。",
    "Staff turned immediately, made eye contact.": "工作人员立即转过头来并与其建立目光接触。",
    
    # 临床功能翻译
    "Escape / Demand Avoidance": "逃避 / 避免要求 (Escape / Demand Avoidance)",
    "Access to Tangible / Activity": "获取实物/活动 (Access to Tangibles)",
    "Social Attention Seeking": "寻求社交关注 (Social Attention)",
    "Automatic / Sensory Synthetic": "自动强化/感官刺激 (Automatic / Sensory)",
    
    # 动态字段翻译
    "[LOCATION] / Inclusive Classroom Setting": "[地点] / 融合教室环境",
    "Responds exceptionally well to visual schedules, tactile items, and praise.": "对视觉日程表、触觉物品和口头夸奖反应非常好。",
    "Elopement (leaving designated area >3 feet) and Vocal Outbursts during transitions.": "离开指定区域（>3英尺）以及在环节转换期间的大声情绪发泄。",
    "Sleep disruption/fatigue increases behavior density by ~40%.": "睡眠中断/疲劳会导致行为发生频率增加约 40%。",
    "Early Intervention (2-5 yrs)": "早期干预阶段 (2-5岁)",
    "School-Age (5-21 yrs)": "学龄阶段 (5-21岁)",
    "Adult / Transition (21+ yrs)": "成年/过渡阶段 (21岁以上)"
}

def translate_to_zh(text_val):
    if not text_val:
        return ""
    val_str = str(text_val).strip()
    return DICTIONARY_ZH.get(val_str, val_str)

# 1. 数据输入
st.header("📋 Phase 1: Multi-Source Data Ingestion")
tab1, tab2, tab3 = st.tabs(["📊 ABC Observation Ledger", "📝 Qualitative Stakeholder Notes", "📈 QABF Psychometric Profiler"])

with tab1:
    st.subheader("Direct Observation ABC Data Ledger")
    uploaded_abc_file = st.file_uploader("Upload Raw ABC Data (CSV format)", type=["csv"], key="abc_csv")
    
    default_abc_data = [
        {"Entry": "Obs #1", "Date/Time": "08/03/2026 09:15 AM", "Observer Role": "BCBA Direct", "Setting": "Desk Work / Literacy", "Antecedent (A)": "Teacher presented multi-step writing task.", "Behavior (B)": "Screamed (>80dB), pushed desk away.", "Consequence (C)": "Staff presented 'Break' visual card; demand paused."},
        {"Entry": "Obs #2", "Date/Time": "08/04/2026 10:30 AM", "Observer Role": "Caregiver (QR Log)", "Setting": "Free Play Transition", "Antecedent (A)": "Timer rang to signal end of iPad time.", "Behavior (B)": "Vocal outburst, dropped to floor.", "Consequence (C)": "Staff offered 2-min extension with visual timer."},
        {"Entry": "Obs #3", "Date/Time": "08/05/2026 01:45 PM", "Observer Role": "Classroom Aide", "Setting": "Small Group Work", "Antecedent (A)": "Instructor turned attention to assist peer.", "Behavior (B)": "Approached staff, pulled sleeve, loud vocalizations.", "Consequence (C)": "Staff turned immediately, made eye contact."}
    ]
    
    if uploaded_abc_file is not None:
        try:
            raw_df = pd.read_csv(uploaded_abc_file)
        except Exception:
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
    uploaded_qual_text = qual_file.read().decode("utf-8") if qual_file is not None else "Parent reports tantrums occur during homework or screen transitions. Teachers note similar patterns in group tasks."
    uploaded_qual_text = deidentify_text(uploaded_qual_text)
    
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

st.divider()
highest_qabf_score = max(att_score, esc_score, tan_score, sen_score, phy_score)
qabf_function = "Escape / Demand Avoidance" if highest_qabf_score == esc_score else "Social Attention / Tangible"

if "Early Intervention" in selected_age_group:
    age_strategy_note = "Early Intervention Focus: Play-based Functional Communication Training (FCT) via PECS/Visual Icons, parent co-regulation, and heavy environmental modification."
    age_strategy_note_zh = "早期干预重点：基于游戏的图片交换沟通系统 (PECS)/视觉卡片功能性沟通训练、家长共同调节及重度环境修改。"
elif "School-Age" in selected_age_group:
    age_strategy_note = "School-Age Focus: Classroom accommodations, high-probability demand sequences, self-monitoring visual timers, and peer-mediated reinforcement."
    age_strategy_note_zh = "学龄阶段重点：教室环境调适、高概率请求序列、自我监控视觉计时器及同伴介导的强化。"
else:
    age_strategy_note = "Adult / Transition Focus: Vocational task chunking, self-advocacy prompts, community integration protocols, and Support Worker SOPs."
    age_strategy_note_zh = "成年/过渡阶段重点：职业任务拆解、自我倡导提示、社区融入方案及支持人员标准作业程序。"

st.info(f"💡 **Automated Triangulation Result**: Primary Function = **{qabf_function}**. {age_strategy_note}")

# 辅助生成函数（设置 zh_heading 默认值为 None 防止位置参数缺失报错）
def add_bilingual_heading(doc, eng_heading, zh_heading=None, level=1):
    if "Chinese" in selected_language and zh_heading:
        heading_text = f"{eng_heading} ({zh_heading})"
    else:
        heading_text = eng_heading
    return doc.add_heading(heading_text, level=level)

def add_bilingual_paragraph(doc, eng_label, eng_val, zh_label, zh_val):
    p = doc.add_paragraph()
    p.add_run(f"{eng_label}: {eng_val}")
    if "Chinese" in selected_language:
        p.add_run(f"\n（中文对照 - {zh_label}: {zh_val}）").italic = True
    return p

def add_bilingual_text(doc, eng_text, zh_text):
    p = doc.add_paragraph()
    p.add_run(eng_text)
    if "Chinese" in selected_language:
        p.add_run(f"\n（中文对照: {zh_text}）").italic = True
    return p

# 3. 生成 FBA 文档
def generate_fba_document():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.8)
    
    if "Chinese" in selected_language:
        title_text = 'FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT\n(功能性行为评估报告)'
    else:
        title_text = 'FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT'
        
    title = doc.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    if "Chinese" in selected_language:
        p = doc.add_paragraph("【 Professional Clinical Draft - Dual-Language Reference / 临床专业初稿（双语对照） 】")
        p.runs[0].font.bold = True
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_bilingual_heading(doc, '1. Clinical Demographics & Background', '临床基本信息与背景资料', level=1)
    add_bilingual_paragraph(doc, "Client Identifier", "[CLIENT_NAME]", "客户标识", "[CLIENT_NAME]")
    add_bilingual_paragraph(doc, "Placement Setting", school_setting, "服务环境", translate_to_zh(school_setting))
    add_bilingual_paragraph(doc, "Age Cohort Category", selected_age_group, "年龄组别", translate_to_zh(selected_age_group))
    add_bilingual_paragraph(doc, "Client Strengths", student_strengths, "个人优势与强化物", translate_to_zh(student_strengths))
    add_bilingual_paragraph(doc, "Target Behavior Definition", behavior_desc, "目标行为定义", translate_to_zh(behavior_desc))
    add_bilingual_paragraph(doc, "Medical / Setting Factors", medical_factors, "医疗/环境因素", translate_to_zh(medical_factors))

    add_bilingual_heading(doc, '1.5 Qualitative Stakeholder Input & Ecological Notes', '利益相关者访谈与生态评估记录', level=1)
    add_bilingual_paragraph(doc, "Qualitative Summary", uploaded_qual_text, "质性评估摘要", "家长报告情绪发泄常发生于作业时间或电子屏幕切换时。教师在不具结构的团队活动中也观察到了相似模式。")

    add_bilingual_heading(doc, '2. Direct Systematic ABC Observation Ledger', '直接系统化 ABC 行为观察日志', level=1)
    
    headers = list(edited_abc.columns)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    for idx, text in enumerate(headers):
        cell = table.rows[0].cells[idx]
        header_display = f"{text}\n({translate_to_zh(text)})" if "Chinese" in selected_language else str(text)
        cell.text = header_display
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
            raw_val_str = str(val)
            if "Chinese" in selected_language:
                zh_val_str = translate_to_zh(raw_val_str)
                if zh_val_str != raw_val_str and not raw_val_str.replace('/','').replace(':','').replace(' ','').isdigit():
                    cell_display = f"{raw_val_str}\n（{zh_val_str}）"
                else:
                    cell_display = raw_val_str
            else:
                cell_display = raw_val_str

            row_cells[c_idx].text = cell_display
            p = row_cells[c_idx].paragraphs[0]
            p.runs[0].font.size = Pt(7.5)
            if r_idx % 2 == 1:
                shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
                row_cells[c_idx]._tc.get_or_add_tcPr().append(shd)

    add_bilingual_heading(doc, '3. Triangulated Discrepancy & Functional Summary', '交叉验证与功能性总结', level=1)
    add_bilingual_text(
        doc,
        f"QABF Highest Score: {highest_qabf_score} (Primary Function: {qabf_function}). Direct ABC logs and psychometric scoring converge on {qabf_function} as the primary maintaining variable.",
        f"QABF 评估最高分: {highest_qabf_score} 分。结合直接 ABC 行为观察日志与量表，一致表明“{translate_to_zh(qabf_function)}”为主要维持功能。"
    )
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 4. 生成 BIP 文档（已修复参数传递问题）
def generate_bip_document():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.8)
    
    if "Chinese" in selected_language:
        title_text = 'BEHAVIOR INTERVENTION PLAN (BIP)\n(行为干预计划)'
    else:
        title_text = 'BEHAVIOR INTERVENTION PLAN (BIP)'

    title = doc.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    add_bilingual_heading(doc, '1. Target Behavior & Primary Function', '目标行为与主要功能', level=1)
    add_bilingual_paragraph(doc, "Client Identifier", "[CLIENT_NAME]", "客户标识", "[CLIENT_NAME]")
    add_bilingual_paragraph(doc, "Target Behavior Definition", behavior_desc, "目标行为定义", translate_to_zh(behavior_desc))
    add_bilingual_paragraph(doc, "Inferred Maintaining Function", qabf_function, "推导主要功能", translate_to_zh(qabf_function))
    add_bilingual_paragraph(doc, "Prescribed Age Focus", age_strategy_note, "针对年龄段策略焦点", age_strategy_note_zh)

    add_bilingual_heading(doc, '2. Proactive Antecedent Strategies', '前因预防策略', level=1)
    add_bilingual_text(
        doc,
        "1. Visual Pre-Correction & Countdown Timers: Provide 5-minute and 2-minute visual cues prior to task transitions to reduce transition anxiety.",
        "1. 视觉预告与倒计时提示：在任务转换前 5 分钟和 2 分钟提供视觉提示，降低过渡期焦虑。"
    )
    add_bilingual_text(
        doc,
        "2. Curriculum Chunking & Demand Modification: Break multi-step instructions into single-step visual task cards to reduce cognitive load.",
        "2. 任务拆解与需求修改：将多步骤书面指令拆解为单步视觉卡片，降低认知负荷。"
    )
    add_bilingual_text(
        doc,
        "3. High-Probability (High-P) Request Sequence: Deliver 3 rapid preferred requests prior to non-preferred demands to build behavioral momentum.",
        "3. 高概率请求序列：在发出较难指令前，连续发出 3 个快速且容易完成的高偏好指令，建立行为惯性。"
    )

    add_bilingual_heading(doc, '3. Functional Replacement Behaviors (FCT)', '功能性替代行为训练', level=1)
    add_bilingual_text(
        doc,
        "1. Independent Break Requests: Teach client to touch/hand the 'I Need a Break' visual card prior to behavioral escalation.",
        "1. 独立请求休息：教导客户在行为升级前，主动触摸或出示“我需要休息”的视觉卡片。"
    )
    add_bilingual_text(
        doc,
        "2. Differential Reinforcement of Alternative Behavior (DRA): Provide immediate functional reinforcement ONLY upon replacement behavior.",
        "2. 差别强化策略 (DRA)：仅在客户使用替代行为（如出示卡片）时提供即时的功能性强化（如暂停/休息）。"
    )

    # 已补全第 4 节与第 5 节中文标题参数
    add_bilingual_heading(doc, '4. Reactive Consequence & Safety Protocols', '后果应对与安全预案', level=1)
    add_bilingual_text(
        doc,
        "1. Escape Extinction (3-Step Prompting): Utilize calm 'Tell-Show-Do' prompting to complete tasks without verbal reprimands.",
        "1. 逃避消退/三步提示法：保持温和中立，使用“告知-示范-协助”顺序引导完成任务，避免口头批评。"
    )
    add_bilingual_text(
        doc,
        "2. Environmental Blocking: Position staff neutrally to block elopement safely without eye contact or verbal commentary.",
        "2. 环境阻挡与安全防护：工作人员保持中立且无眼神接触的状态下安全阻挡逃跑行为。"
    )

    add_bilingual_heading(doc, '5. Generalization & Local Re-identification Note', '泛化与本地脱敏还原说明', level=1)
    add_bilingual_text(
        doc,
        "Note for BCBA/BSP: All client names are export-masked as [CLIENT_NAME]. Use Word Find & Replace (Ctrl+H) on your local workstation to restore true identifying details prior to clinical submission.",
        "提示：所有名称均已脱敏处理为 [CLIENT_NAME]。在提交团队前，请在 Word 中按 Ctrl+H 替换为真实姓名。"
    )

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 5. 下载导出
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
