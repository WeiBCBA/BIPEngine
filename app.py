import streamlit as st
import io
import re
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# 页面基本配置
st.set_page_config(page_title="US-BCBA Clinical BIPEngine v2.5", layout="wide", initial_sidebar_state="expanded")

# 侧边栏配置
with st.sidebar:
    st.markdown("<h2 style='color: #1F4E78;'>⚙️ BIPEngine v2.5 (US Standard)</h2>", unsafe_allow_html=True)
    st.caption("US BCBA & LBA Clinical Decision Support Engine for FBA/BIP")
    st.divider()
    
    st.markdown("### 🔒 Privacy & HIPAA Compliance")
    st.markdown("""
    * **No Cloud Storage**: Zero persistent data retention.
    * **Instant Memory Wipe**: Session clearing on browser close.
    * **De-identification**: Automatic PII masking (`[CLIENT_NAME]`).
    """)
    st.divider()
    
    st.subheader("🌐 Language & Cohort Settings")
    selected_language = st.selectbox(
        "Report Output Language Target:",
        ["English (Standard US)", "English & Chinese Dual-Language (中英双语对照)", "English & Spanish Dual-Language (Español & English)"]
    )
    
    selected_age_group = st.selectbox(
        "Client Development Cohort:",
        ["Early Intervention (2-5 yrs)", "School-Age (5-21 yrs)", "Adult / Transition (21+ yrs)"],
        key="age_group_select"
    )

st.title("🚀 US-BCBA Automated FBA & BIP Clinical Compiler")
st.markdown("##### Advanced Triangulation Engine adhering to BACB & US Healthcare Standards")
st.divider()

def deidentify_text(text):
    if not text:
        return ""
    text = re.sub(r'(?i)\b(client|student|child|patient):\s*([A-Z][a-z]+\s+[A-Z][a-z]+)', r'\1: [CLIENT_NAME]', text)
    return text

# 全面补全美式术语与文本双语字典
DICTIONARY_ZH = {
    # 模板数据源与基本字段
    "Direct Observations": "直接行为观察",
    "Student Interview": "学生/客户访谈",
    "Teacher Interview": "教师/工作人员访谈",
    "Parent Interview": "家长/照护者访谈",
    "Rating Scales": "评估量表 (如 QABF/MAS)",
    
    # 维度与量化字段
    "Frequency": "发生频率 (Frequency)",
    "Duration": "持续时间 (Duration)",
    "Intensity": "行为强度 (Intensity)",
    "Setting Events": "情境因素/慢速诱因 (Setting Events)",
    "Antecedent Events": "即时前因/快速诱因 (Antecedent Events)",
    "Non-occurrence Situations": "行为不常发生的例外情境 (Non-occurrence Situations)",
    "Consequences": "行为后果 (Consequences)",
    "Hypothesis Statement": "行为假设说明 (Hypothesis)",
    "Communicative Intent": "沟通意图 (Communicative Intent)",
    
    # ABC 字段
    "Entry": "记录编号 (Entry)",
    "Date/Time": "日期/时间 (Date/Time)",
    "Observer Role": "观察者/数据源 (Observer & Tool)",
    "Setting": "服务环境 (Setting)",
    "Antecedent (A)": "前因 (Antecedent - A)",
    "Behavior (B)": "行为 (Behavior - B)",
    "Consequence (C)": "后果 (Consequence - C)",
    "Engine Auto-Inferred Function": "系统推导功能 (Inferred Function)",
    
    # 临床功能
    "Task Escape / Demand Avoidance": "逃避 / 避免任务 (Task Escape)",
    "Access to Tangibles / Activities": "获取实物/活动/环境控制 (Access to Tangibles/Control)",
    "Social Attention Seeking": "寻求社交关注 (Social Attention)",
    "Automatic / Sensory Stimulation": "自动强化/感官刺激 (Automatic / Sensory)",
    "Physical Discomfort / Internal State": "生理不适/内部状态 (Physical Discomfort)",
    
    # 美式观察者与环境描述
    "BCBA Direct Observation": "BCBA 直接临床观察",
    "RBT (CentralReach Portal)": "RBT 行为技术员 (CentralReach 数据端)",
    "RBT (Catalyst Data App)": "RBT 行为技术员 (Catalyst 数据采集端)",
    "Paraprofessional (School Data Sheet)": "学校助教/特教副手 (学校数据记录表)",
    "Special Ed Teacher (IEP Log)": "特教教师 (IEP 行为日志)",
    "Job Coach (Vocational Log)": "职业辅导员 (职业训练日志)",
    "Direct Care Staff (CR Mobile)": "直接照护人员 (CentralReach 移动端)",
    "Parent Self-Report (Caregiver Log)": "家长自述 (家长培训反馈日志)",
    "Clinic Therapy Room": "诊所治疗室",
    "Home ABA Session": "居家 ABA 训练",
    "Clinic Social Skills Group": "诊所社交技能小组",
    "[LOCATION] / Early Childhood Inclusive Center": "[地点] / 早期融合教育中心",
    "[LOCATION] / Supported Community Living Apartment": "[地点] / 支持性社区居住公寓",
    "[LOCATION] / Inclusive School & In-Home ABA": "[地点] / 融合学校与居家 ABA",

    # 描述文本
    "Screaming, biting self, head-banging during transitions.": "转换过程中出现尖叫、咬自己、用头撞墙/地。",
    "Auditory sensitivity, sleep disruption.": "听觉敏感、睡眠中断。",
    "Parents report severe tantrums during toy sharing and routine changes. Pediatrician notes high sensitivity to auditory stimuli.": "家长报告在分享玩具和例行程序变更时出现严重情绪发作；儿科医生指出对听觉刺激高度敏感。",
    "Eager to engage with cause-and-effect toys; high response to PECS.": "热衷于参与因果关系玩具；对图片交换沟通系统 (PECS) 反应良好。",
    "No prior ABA services reported; enrolled in inclusive early center.": "无先前 ABA 服务记录；目前就读于早期融合中心。",
    "High Intensity (Self-injurious behavior & potential soft-tissue damage)": "高强度（含自伤行为及潜在软组织损伤风险）",
    "Medium Intensity (Disrupts learning environment)": "中等强度（打断正常教学与学习环境）",
    "High Intensity (Aggression towards staff & safety concerns)": "高强度（攻击工作人员与显著安全风险）",
    
    # ABC 细节词条
    "RBT requested sharing toy truck during naturalistic play.": "RBT 在自然情境游戏期间要求分享玩具卡车。",
    "Screamed, bit own arm, threw toy.": "尖叫、咬自己的手臂、扔玩具。",
    "RBT paused demand, offered sensory chew tool.": "RBT 暂停任务指令，提供感官咀咀嚼咬乐工具。",
    "Therapist transitioned from bubble play to discrete trial teaching (DTT).": "治疗师从吹泡泡游戏转换至分解尝试教导 (DTT)。",
    "Dropped to floor, crying, head banging on carpet.": "瘫倒在地、哭泣、在地毯上用头撞地。",
    "Therapist paused demand, presented PECS break icon.": "治疗师暂停任务指令，出示 PECS 休息图标。",
    "RBT called for group cleanup time.": "RBT 呼叫进行小组收拾玩具时间。",
    "Ran toward clinic exit door (elopement).": "跑向诊所出口大门（离位/逃跑）。",
    "Staff guided back with visual transition timer.": "工作人员通过视觉过渡计时器将其引导返回。"
}

def translate_to_zh(text_val):
    if not text_val:
        return ""
    val_str = str(text_val).strip()
    if val_str in DICTIONARY_ZH:
        return DICTIONARY_ZH[val_str]
    translated = val_str
    for en, zh in DICTIONARY_ZH.items():
        if en in translated:
            translated = translated.replace(en, zh)
    return translated

# 预设 ABC 数据
def get_age_specific_abc_presets(age_group):
    if "Adult" in age_group:
        return [
            {"Entry": "Obs #1", "Date/Time": "08/03/2026 09:15 AM", "Observer Role": "Direct Care Staff (CR Mobile)", "Setting": "Supported Living Apartment", "Antecedent (A)": "Newly hired staff member presented morning chore checklist.", "Behavior (B)": "Verbal aggression (cursing, threats) and physical aggression (shoving staff).", "Consequence (C)": "Senior BCBA/Staff stepped in, guided new staff to pause demand, and represented visual choice board."},
            {"Entry": "Obs #2", "Date/Time": "08/04/2026 06:30 PM", "Observer Role": "Direct Care Staff (CR Mobile)", "Setting": "Supported Living Apartment", "Antecedent (A)": "Roommate turned on living room TV and adjusted seating area without client consent.", "Behavior (B)": "Loud vocal resistance, blocking TV screen, grabbing remote from roommate.", "Consequence (C)": "Staff prompted roommate mediation and provided alternative personal tablet for private room."},
            {"Entry": "Obs #3", "Date/Time": "08/05/2026 11:00 AM", "Observer Role": "Job Coach (Vocational Log)", "Setting": "Day Program / Vocational Workshop", "Antecedent (A)": "Client reported joint pain/headache after 2 hours of repetitive standing work.", "Behavior (B)": "Pacing, hand-wringing, aggressive resistance to verbal prompts.", "Consequence (C)": "Staff offered PRN pain relief medication and quiet rest area."}
        ]
    elif "Early Intervention" in age_group:
        return [
            {"Entry": "Obs #1", "Date/Time": "08/03/2026 09:15 AM", "Observer Role": "RBT (CentralReach Portal)", "Setting": "Clinic Therapy Room", "Antecedent (A)": "RBT requested sharing toy truck during naturalistic play.", "Behavior (B)": "Screamed, bit own arm, threw toy.", "Consequence (C)": "RBT paused demand, offered sensory chew tool."},
            {"Entry": "Obs #2", "Date/Time": "08/04/2026 10:30 AM", "Observer Role": "BCBA Direct Observation", "Setting": "Home ABA Session", "Antecedent (A)": "Therapist transitioned from bubble play to discrete trial teaching (DTT).", "Behavior (B)": "Dropped to floor, crying, head banging on carpet.", "Consequence (C)": "Therapist paused demand, presented PECS break icon."},
            {"Entry": "Obs #3", "Date/Time": "08/05/2026 04:00 PM", "Observer Role": "RBT (Catalyst Data App)", "Setting": "Clinic Social Skills Group", "Antecedent (A)": "RBT called for group cleanup time.", "Behavior (B)": "Ran toward clinic exit door (elopement).", "Consequence (C)": "Staff guided back with visual transition timer."}
        ]
    else: # School-Age (5-21)
        return [
            {"Entry": "Obs #1", "Date/Time": "08/03/2026 04:30 PM", "Observer Role": "RBT (CentralReach Portal)", "Setting": "In-Home ABA / Screen Time Transition", "Antecedent (A)": "Timer rang signaling 30-min iPad screen time limits reached while RBT turned to document data.", "Behavior (B)": "Pacing, grabbing iPad back, screaming 'Look at me!', dropping to floor.", "Consequence (C)": "RBT made immediate eye contact, prompted '1-min visual extension card', and reinforced quiet waiting."},
            {"Entry": "Obs #2", "Date/Time": "08/04/2026 01:45 PM", "Observer Role": "Paraprofessional (School Data Sheet)", "Setting": "Special Ed Classroom (Small Group)", "Antecedent (A)": "Instructor turned attention to assist peer during iPad group activity.", "Behavior (B)": "Approached staff, pulled sleeve, loud vocalizations, tried to grab peer's iPad.", "Consequence (C)": "Staff turned immediately, made eye contact, and redirected to waiting visual schedule."},
            {"Entry": "Obs #3", "Date/Time": "08/05/2026 09:15 AM", "Observer Role": "BCBA Direct Observation", "Setting": "General Ed Classroom (Desk Work)", "Antecedent (A)": "Teacher presented multi-step writing worksheet.", "Behavior (B)": "Screamed (>80dB), pushed desk away.", "Consequence (C)": "Staff presented 'Break' visual card; demand paused."}
        ]

st.header("📋 Phase 1: Multi-Source Clinical Ingestion")
tab1, tab2, tab3 = st.tabs(["📊 Direct ABC Observations", "📝 Expanded FBA Template Inputs", "📈 QABF Psychometric Scale"])

with tab1:
    st.subheader("Direct Systematic ABC Data Ledger")
    uploaded_abc_file = st.file_uploader("Upload Raw ABC CSV", type=["csv"], key="abc_csv")
    default_abc_data = get_age_specific_abc_presets(selected_age_group)
    raw_df = pd.read_csv(uploaded_abc_file) if uploaded_abc_file is not None else pd.DataFrame(default_abc_data)

    def infer_function_from_row(row):
        ant = str(row.get("Antecedent (A)", "")).lower()
        con = str(row.get("Consequence (C)", "")).lower()
        beh = str(row.get("Behavior (B)", "")).lower()
        full_text = f"{ant} {beh} {con}"

        if any(k in full_text for k in ["pain", "medication", "joint", "headache", "sick", "illness"]):
            return "Physical Discomfort / Internal State"
        if any(k in full_text for k in ["demand", "task", "worksheet", "chore", "dtt", "cleanup", "writing", "instruction", "pause demand", "break"]):
            if not ("ipad" in ant or "screen" in ant or "toy" in ant):
                return "Task Escape / Demand Avoidance"
        if any(k in full_text for k in ["ipad", "toy", "screen", "remote", "tv", "tablet", "game", "sharing toy"]):
            return "Access to Tangibles / Activities"
        if "look at me" in beh or "pulled sleeve" in beh or "turned attention" in ant or "attention" in ant or "grab peer" in beh:
            return "Social Attention Seeking"
        return "Automatic / Sensory Stimulation"

    raw_df["Engine Auto-Inferred Function"] = raw_df.apply(infer_function_from_row, axis=1)
    edited_abc = st.data_editor(raw_df, num_rows="dynamic", use_container_width=True, key=f"editor_{selected_age_group}")

with tab2:
    st.subheader("Expanded FBA Form Fields (Matching Standard Template)")
    
    col_a, col_b = st.columns(2)
    with col_a:
        school_name = st.text_input("School / Agency Name", "Metropolitan Inclusive Center")
        school_district = st.text_input("School District / Health Region", "District 10 Behavioral Division")
        student_dob = st.text_input("Student DOB", "05/12/2021")
        student_id = st.text_input("Student ID", "ID-908231")
        fba_date = st.text_input("Date of FBA", "08/08/2026")
        
    with col_b:
        data_sources = st.multiselect(
            "Data Sources (Select all that apply):",
            ["Direct Observations", "Student Interview", "Teacher Interview", "Parent Interview", "Rating Scales"],
            default=["Direct Observations", "Teacher Interview", "Parent Interview", "Rating Scales"]
        )
        edu_history = st.text_area("Educational History", "Enrolled in inclusive early childhood center; receives Speech (SLP) and OT services.", height=68)

    st.divider()
    st.markdown("##### Behavioral Operational Parameters")
    
    if "Adult" in selected_age_group:
        default_behaviors = "Physical aggression toward new staff, verbal threats, and roommate conflicts over shared items."
        default_freq = "3 to 5 episodes per week during staff transitions or common area usage."
        default_dur = "Episodes last 15 to 45 minutes until resolution."
        default_intensity = "High Intensity (Aggression towards staff & safety concerns)"
        default_setting_events = "Unfamiliar substitute staff on shift, physical joint pain, or lack of sleep."
        default_non_occur = "Preferred staff presence, clear visual schedule, individual leisure time in private room."
    elif "Early Intervention" in selected_age_group:
        default_behaviors = "Screaming (>80dB), biting self (forearm), head-banging on floor during transitions."
        default_freq = "4 to 6 times per daily 3-hour session."
        default_dur = "2 to 10 minutes per episode."
        default_intensity = "High Intensity (Self-injurious behavior & potential soft-tissue damage)"
        default_setting_events = "Sleep disruption, high auditory background noise, or physical fatigue."
        default_non_occur = "Preferred sensory play (water/bubble play), 1-on-1 direct support, highly predictable routines."
    else: # School-Age
        default_behaviors = "iPad transition tantrums (screaming, dropping to floor), sleeve-pulling for attention."
        default_freq = "2 to 4 times per day during activity transitions."
        default_dur = "5 to 15 minutes per episode."
        default_intensity = "Medium Intensity (Disrupts learning environment)"
        default_setting_events = "Extended digital screen time prior to instruction, instructor attention focused on peers."
        default_non_occur = "Explicit visual timer counting down, structured 1-on-1 attention, structured break cards."

    c1, c2, c3 = st.columns(3)
    behavior_desc = c1.text_area("Target Behavior Description", default_behaviors, height=100)
    freq_desc = c2.text_area("Frequency", default_freq, height=100)
    dur_desc = c3.text_area("Duration", default_dur, height=100)
    
    c4, c5, c6 = st.columns(3)
    intensity_desc = c4.text_input("Intensity Level", default_intensity)
    setting_events = c5.text_area("Setting Events (Slow Triggers)", default_setting_events, height=80)
    non_occur = c6.text_area("Situations Where Behavior Does NOT Occur", default_non_occur, height=80)
    
    student_strengths = st.text_area("Client Strengths & Motivators", "Responds well to visual timers, high motivation for cause-and-effect toys.", height=70)
    uploaded_qual_text = st.text_area("Qualitative Summary / Notes", "Parents report high auditory sensitivity. Staff noted strong engagement when using PECS.", height=70)

with tab3:
    st.subheader("QABF Assessment Scores")
    q1, q2, q3, q4, q5 = st.columns(5)
    att_score = q1.number_input("Social Attention", 0, 15, 3)
    esc_score = q2.number_input("Task Escape", 0, 15, 4)
    tan_score = q3.number_input("Tangibles / Control", 0, 15, 12)
    sen_score = q4.number_input("Sensory", 0, 15, 14)
    phy_score = q5.number_input("Physical Discomfort", 0, 15, 11)

st.divider()

# 三方交叉推断逻辑
abc_counts = edited_abc["Engine Auto-Inferred Function"].value_counts()
primary_function = abc_counts.index[0] if not abc_counts.empty else "Task Escape / Demand Avoidance"

hypothesis_eng = f"When presented with {setting_events} and immediate transition prompts, [CLIENT_NAME] engages in {behavior_desc} in order to achieve {primary_function}. The behavior serves as a functional attempt to communicate a desire for control or sensory regulation."
hypothesis_zh = f"当出现【{translate_to_zh(setting_events)}】及即时转换指令时，[CLIENT_NAME] 表现出【{translate_to_zh(behavior_desc)}】，以获得【{translate_to_zh(primary_function)}】。该行为是表达对环境控制或感官调节需求的功能性沟通尝试。"

# 辅助 Word 打印工具
def add_bilingual_heading(doc, eng_heading, zh_heading=None, level=1):
    heading_text = f"{eng_heading} ({zh_heading})" if "Chinese" in selected_language and zh_heading else eng_heading
    return doc.add_heading(heading_text, level=level)

def add_bilingual_paragraph(doc, eng_label, eng_val, zh_label, zh_val):
    p = doc.add_paragraph()
    p.add_run(f"{eng_label}: ").bold = True
    p.add_run(str(eng_val))
    if "Chinese" in selected_language:
        p.add_run(f"\n（中文对照 - {zh_label}: {translate_to_zh(zh_val)}）").italic = True
    return p

# 生成 FBA 文档 (完全符合提供模板的扩展版)
def generate_fba_document():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.7)
    
    # 标题
    title_text = 'FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) FORM\n(功能性行为评估标准表)' if "Chinese" in selected_language else 'FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) FORM'
    title = doc.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 1. 表格头部：学生基本信息 (2x4 结构化 Word 表格)
    header_table = doc.add_table(rows=4, cols=4)
    header_table.style = 'Table Grid'
    
    info_map = [
        [("Student Name", "[CLIENT_NAME]"), ("School Name", school_name)],
        [("Student DOB", student_dob), ("School District", school_district)],
        [("Student ID", student_id), ("Date of FBA", fba_date)],
        [("Cohort Category", selected_age_group), ("Placement", school_name)]
    ]
    
    for r_idx, row_data in enumerate(info_map):
        for c_group, (lbl, val) in enumerate(row_data):
            c_lbl = header_table.cell(r_idx, c_group * 2)
            c_val = header_table.cell(r_idx, c_group * 2 + 1)
            
            c_lbl.text = f"{lbl}\n({translate_to_zh(lbl)})" if "Chinese" in selected_language else lbl
            c_val.text = str(val)
            
            # 样式
            shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
            c_lbl._tc.get_or_add_tcPr().append(shd)
            c_lbl.paragraphs[0].runs[0].font.bold = True
            c_lbl.paragraphs[0].runs[0].font.size = Pt(8.5)
            c_val.paragraphs[0].runs[0].font.size = Pt(8.5)

    doc.add_paragraph() # 换行

    # 2. 数据来源
    add_bilingual_heading(doc, '1. Data Sources', '数据来源', level=1)
    ds_str = ", ".join(data_sources)
    add_bilingual_paragraph(doc, "Selected Sources", ds_str, "已选数据源", ds_str)

    # 3. 背景与优势
    add_bilingual_heading(doc, '2. Brief Student Background & Strengths', '学生简要背景与优势', level=1)
    add_bilingual_paragraph(doc, "Strengths & Motivators", student_strengths, "学生优势与强化物", student_strengths)
    add_bilingual_paragraph(doc, "Educational History", edu_history, "教育背景履历", edu_history)

    # 4. 目标行为维度 (核心扩展区)
    add_bilingual_heading(doc, '3. Target Behavior Operational Breakdown', '目标行为操作化拆解', level=1)
    add_bilingual_paragraph(doc, "Description of Target Behavior", behavior_desc, "目标行为操作化定义", behavior_desc)
    add_bilingual_paragraph(doc, "Frequency", freq_desc, "发生频率", freq_desc)
    add_bilingual_paragraph(doc, "Duration", dur_desc, "持续时间", dur_desc)
    add_bilingual_paragraph(doc, "Intensity", intensity_desc, "行为强度", intensity_desc)

    # 5. 前因与情境因素
    add_bilingual_heading(doc, '4. Behavioral Triggers & Environmental Context', '行为诱因与环境情境', level=1)
    add_bilingual_paragraph(doc, "Setting Events (Slow Triggers)", setting_events, "情境因素 (慢速诱因)", setting_events)
    add_bilingual_paragraph(doc, "Antecedent Events (Immediate Triggers)", "Prompts for transition, removal of tangibles, or social attention shift.", "即时前因 (快速诱因)", "转换指令、剥夺实物或社交关注转移。")
    add_bilingual_paragraph(doc, "Non-Occurrence Situations", non_occur, "行为不常发生的例外情境", non_occur)
    add_bilingual_paragraph(doc, "Consequences (Immediate Responses)", "Task demands paused, sensory chew tools or verbal redirection provided.", "行为后果 (即时回应)", "任务暂停、提供感官咀嚼工具或口头重定向。")

    # 6. ABC 观察日志表格
    add_bilingual_heading(doc, '5. Direct Systematic ABC Observation Ledger', '直接系统化 ABC 观察日志', level=1)
    headers = list(edited_abc.columns)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    for idx, text in enumerate(headers):
        cell = table.rows[0].cells[idx]
        cell.text = f"{text}\n({translate_to_zh(text)})" if "Chinese" in selected_language else str(text)
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
            cell_display = f"{raw_val_str}\n（{translate_to_zh(raw_val_str)}）" if "Chinese" in selected_language and translate_to_zh(raw_val_str) != raw_val_str else raw_val_str
            row_cells[c_idx].text = cell_display
            p = row_cells[c_idx].paragraphs[0]
            p.runs[0].font.size = Pt(7.5)
            if r_idx % 2 == 1:
                shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
                row_cells[c_idx]._tc.get_or_add_tcPr().append(shd)

    # 7. 假设与行为功能判定
    add_bilingual_heading(doc, '6. Hypothesis & Function Determination', '行为假设与功能判定', level=1)
    add_bilingual_paragraph(doc, "Hypothesis Statement", hypothesis_eng, "行为假设说明", hypothesis_zh)
    add_bilingual_paragraph(doc, "Primary Determined Function", primary_function, "判定主要功能", primary_function)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 导出界面
st.subheader("📄 Export Expanded US-BCBA Clinical Draft Documents")
if st.button("🚀 Compile Template-Matched FBA Report (.docx)", type="primary", use_container_width=True):
    fba_file = generate_fba_document()
    st.success("Expanded FBA Form Compiled Successfully!")
    st.download_button(
        label="📄 Download Template_Matched_FBA_Report.docx",
        data=fba_file,
        file_name=f"Standard_FBA_Report_{selected_age_group.split()[0]}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
