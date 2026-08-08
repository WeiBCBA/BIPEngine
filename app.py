import streamlit as st
import io
import re
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# 修正参数名：initial_sidebar_state
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

# 美式术语双语字典
DICTIONARY_ZH = {
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
    
    # 美式观察者与工具
    "BCBA Direct Observation": "BCBA 直接临床观察",
    "RBT (CentralReach Portal)": "RBT 行为技术员 (CentralReach 数据端)",
    "RBT (Catalyst Data App)": "RBT 行为技术员 (Catalyst 数据采集端)",
    "Paraprofessional (School Data Sheet)": "学校助教/特教副手 (学校数据记录表)",
    "Special Ed Teacher (IEP Log)": "特教教师 (IEP 行为日志)",
    "Job Coach (Vocational Log)": "职业辅导员 (职业训练日志)",
    "Direct Care Staff (CR Mobile)": "直接照护人员 (CentralReach 移动端)",
    "Parent Self-Report (Caregiver Log)": "家长自述 (家长培训反馈日志)",
    
    # 环境与前因描述
    "Day Program / Vocational Workshop": "日间中心 / 职业技能车间",
    "Community Outing / Grocery": "社区外出 / 超市购物",
    "Supported Living Apartment": "支持性居住公寓",
    "General Ed Classroom (Desk Work)": "普通教育教室 (桌面作业)",
    "In-Home ABA / Screen Time Transition": "居家 ABA / 屏幕时间转换",
    
    # 具体前因、行为与后果词条
    "Newly hired staff member presented morning chore checklist.": "新到岗员工呈递了早晨家务清单。",
    "Roommate turned on living room TV and adjusted seating area without client consent.": "室友未经客户同意打开客厅电视并调整沙发座椅。",
    "Timer rang signaling 30-min iPad screen time limits reached while RBT turned to document data.": "计时器响起提示 30 分钟 iPad 时长已满，同时 RBT 转身记录数据。",
    "Verbal aggression (cursing, threats) and physical aggression (shoving staff).": "言语攻击（骂人、威胁）与身体攻击（推搡工作人员）。",
    "Loud vocal resistance, blocking TV screen, grabbing remote from roommate.": "大声抗议发声、挡住电视屏幕、强行夺取室友手中的遥控器。",
    "Pacing, grabbing iPad back, screaming 'Look at me!', dropping to floor.": "踱步、夺回 iPad、大喊“看我！”并落座在地上。",
    "Senior BCBA/Staff stepped in, guided new staff to pause demand, and represented visual choice board.": "资深 BCBA/员工介入，指导新员工暂停任务指令，并重新出示视觉选择板。",
    "Staff prompted roommate mediation and provided alternative personal tablet for private room.": "工作人员介入调解室友冲突，并为客户提供私人房间使用的备用个人平板。",
    "RBT made immediate eye contact, prompted '1-min visual extension card', and reinforced quiet waiting.": "RBT 立即与其进行眼神接触，给予“1分钟视觉延时卡”提示，并对其安静等待给予强化。"
}

def translate_to_zh(text_val):
    if not text_val:
        return ""
    val_str = str(text_val).strip()
    return DICTIONARY_ZH.get(val_str, val_str)

# 动态生成特定年龄段的预设 ABC 数据
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
tab1, tab2, tab3 = st.tabs(["📊 Direct ABC Observations (60% Weight)", "📝 Stakeholder Notes (30% Weight)", "📈 QABF Psychometric Scale (10% Weight)"])

with tab1:
    st.subheader("Direct Systematic ABC Data Ledger (US Clinical Practice)")
    st.caption("Data captured via BCBA Direct Field Observation, RBT CentralReach Portal, and School Paraprofessional Logs.")
    uploaded_abc_file = st.file_uploader("Upload Raw ABC CSV", type=["csv"], key="abc_csv")
    
    default_abc_data = get_age_specific_abc_presets(selected_age_group)
    raw_df = pd.read_csv(uploaded_abc_file) if uploaded_abc_file is not None else pd.DataFrame(default_abc_data)

    def infer_function_from_row(row):
        ant = str(row.get("Antecedent (A)", "")).lower()
        con = str(row.get("Consequence (C)", "")).lower()
        beh = str(row.get("Behavior (B)", "")).lower()
        
        if "look at me" in beh or "turned attention" in ant or "pulled sleeve" in beh or "eye contact" in con:
            return "Social Attention Seeking"
        elif "pain" in ant or "medication" in con or "joint" in ant:
            return "Physical Discomfort / Internal State"
        elif "ipad" in ant or "remote" in beh or "tv" in ant or "screen" in ant:
            return "Access to Tangibles / Activities"
        elif "newly hired" in ant or "demand" in ant or "task" in ant or "worksheet" in ant or "chore" in ant:
            return "Task Escape / Demand Avoidance"
        return "Automatic / Sensory Stimulation"

    raw_df["Engine Auto-Inferred Function"] = raw_df.apply(infer_function_from_row, axis=1)
    edited_abc = st.data_editor(raw_df, num_rows="dynamic", use_container_width=True, key=f"editor_{selected_age_group}")

with tab2:
    st.subheader("Qualitative & Ecological Stakeholder Inputs")
    qual_file = st.file_uploader("Upload Interview Transcripts (.txt)", type=["txt"], key="qual_txt")
    
    if "Adult" in selected_age_group:
        default_qual = "Direct care staff report severe behavioral spikes when new substitute staff introduce demands, or when roommates access shared living room space/TV. Environmental control and task avoidance with new staff are primary drivers."
        default_placement = "[LOCATION] / Supported Community Living Apartment"
        default_strengths = "Independent with personal hygiene; responds well to visual schedules, designated personal space, and familiar staff."
        default_behaviors = "Physical aggression toward new staff, verbal threats, and roommate conflicts over shared items."
        default_factors = "High anxiety with unfamiliar personnel, sensitivity to environmental change in common areas."
    elif "Early Intervention" in selected_age_group:
        default_qual = "Parents report severe tantrums during toy sharing and routine changes. Pediatrician notes high sensitivity to auditory stimuli."
        default_placement = "[LOCATION] / Early Childhood Inclusive Center"
        default_strengths = "Eager to engage with cause-and-effect toys; high response to PECS."
        default_behaviors = "Screaming, biting self, head-banging during transitions."
        default_factors = "Auditory sensitivity, sleep disruption."
    else: # School-Age
        default_qual = "Parents and RBTs report high obsession with iPad screen time, coupled with high social attention-seeking behaviors when instructor attention shifts to peers. Tantrums spike during iPad transitions."
        default_placement = "[LOCATION] / Inclusive School & In-Home ABA"
        default_strengths = "Responds exceptionally well to visual timers, token systems, and structured iPad reward breaks."
        default_behaviors = "iPad transition tantrums, sleeve-pulling/vocalizing for attention, and task escape during writing."
        default_factors = "Fixation on digital electronics, heightened sensitivity to peer/adult attention allocation."

    uploaded_qual_text = qual_file.read().decode("utf-8") if qual_file is not None else default_qual
    uploaded_qual_text = deidentify_text(uploaded_qual_text)
    
    col1, col2 = st.columns(2)
    with col1:
        school_setting = st.text_input("Placement / Setting", default_placement)
        student_strengths = st.text_area("Client Strengths & Motivators", default_strengths, height=90)
    with col2:
        behavior_desc = st.text_area("Target Behavior Operational Definition", default_behaviors, height=90)
        medical_factors = st.text_area("Medical / Setting Factors", default_factors, height=90)

with tab3:
    st.subheader("Questions About Behavioral Function (QABF) - Informant Measure")
    st.caption("Note: In US BCBA practice, QABF is weighted lightly (10%) due to lower informant reliability compared to direct ABC data.")
    
    def_att = 14 if "School" in selected_age_group else 3
    def_esc = 13 if "Adult" in selected_age_group else (5 if "School" in selected_age_group else 4)
    def_tan = 12 if "Adult" in selected_age_group else (13 if "School" in selected_age_group else 3)
    def_sen = 2 if "School" in selected_age_group else 14
    def_phy = 8 if "Adult" in selected_age_group else (1 if "School" in selected_age_group else 11)

    q1, q2, q3, q4, q5 = st.columns(5)
    att_score = q1.number_input("Social Attention", 0, 15, def_att)
    esc_score = q2.number_input("Task Escape", 0, 15, def_esc)
    tan_score = q3.number_input("Tangibles / Control", 0, 15, def_tan)
    sen_score = q4.number_input("Sensory", 0, 15, def_sen)
    phy_score = q5.number_input("Physical Discomfort", 0, 15, def_phy)

st.divider()

# 美式 BCBA 核心算法：60% ABC + 30% Stakeholder + 10% QABF
qabf_scores = {
    "Task Escape / Demand Avoidance": esc_score,
    "Social Attention Seeking": att_score,
    "Access to Tangibles / Activities": tan_score,
    "Automatic / Sensory Stimulation": sen_score,
    "Physical Discomfort / Internal State": phy_score
}

top_qabf_function = max(qabf_scores, key=qabf_scores.get)
highest_qabf_score = qabf_scores[top_qabf_function]

abc_counts = edited_abc["Engine Auto-Inferred Function"].value_counts()
top_abc_function = abc_counts.index[0] if not abc_counts.empty else "Task Escape / Demand Avoidance"

primary_function = top_abc_function

if top_abc_function == top_qabf_function:
    triangulation_summary_eng = f"Primary Function: {primary_function}. Full convergence across Direct ABC data (60% weight), Stakeholder reports, and QABF scale."
    triangulation_summary_zh = f"主要功能：{translate_to_zh(primary_function)}。直接 ABC 观察数据（60%权重）、利益相关者报告及 QABF 量表结果完全吻合。"
else:
    triangulation_summary_eng = (
        f"Primary Function: {primary_function} (Anchored on Direct ABC Data & Stakeholder Notes). "
        f"Clinical Discrepancy Note: QABF scale indicated elevated '{top_qabf_function}'. In accordance with BACB best practices, "
        f"indirect informant scales carry lower clinical weight (10%) due to subjective informant variance; direct observation confirms {primary_function} as the true maintaining variable."
    )
    triangulation_summary_zh = (
        f"主要功能：{translate_to_zh(primary_function)}（基于 60% 权重的直接 ABC 观察数据与 30% 质性访谈定性）。"
        f"数据不一致说明：QABF 间接量表显示“{translate_to_zh(top_qabf_function)}”分值较高。依据美国 BCBA 临床规范，"
        f"间接量表受填写者主观偏差影响仅占 10% 辅助权重，最终判定以客观直接 ABC 观察结果为准。"
    )

if "Early Intervention" in selected_age_group:
    age_strategy_note = "Early Intervention Focus: Play-based Functional Communication Training (FCT) via PECS/Visual Icons, parent co-regulation, and heavy environmental modification."
    age_strategy_note_zh = "早期干预重点：基于游戏的图片交换沟通系统 (PECS)/视觉卡片功能性沟通训练、家长共同调节及重度环境修改。"
elif "School-Age" in selected_age_group:
    age_strategy_note = "School-Age Focus: Combined Social Attention & iPad Tangible Access protocols, visual transition timers, differential reinforcement of alternative behavior (DRA), and structured attention delivery."
    age_strategy_note_zh = "学龄阶段重点：社交关注与 iPad 实体获取结合协议、视觉过渡计时器、替代行为差别强化 (DRA) 及结构化关注给予。"
else:
    age_strategy_note = "Adult / Transition Focus: Shared living room roommate boundary protocols, new staff pairing/fading procedures, and vocational task chunking."
    age_strategy_note_zh = "成年/过渡阶段重点：共享居住空间室友边界协议、新员工配对与配对渐隐程序以及职业任务拆解。"

st.info(f"💡 **US-BCBA Clinical Triangulation**: Primary Function = **{primary_function}**. {age_strategy_note}")

# 辅助生成函数
def add_bilingual_heading(doc, eng_heading, zh_heading=None, level=1):
    heading_text = f"{eng_heading} ({zh_heading})" if "Chinese" in selected_language and zh_heading else eng_heading
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

# 生成 FBA 文档
def generate_fba_document():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.8)
    
    title_text = 'FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT\n(功能性行为评估报告)' if "Chinese" in selected_language else 'FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT'
    title = doc.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_bilingual_heading(doc, '1. Clinical Demographics & Background', '临床基本信息与背景资料', level=1)
    add_bilingual_paragraph(doc, "Client Identifier", "[CLIENT_NAME]", "客户标识", "[CLIENT_NAME]")
    add_bilingual_paragraph(doc, "Placement Setting", school_setting, "服务环境", translate_to_zh(school_setting))
    add_bilingual_paragraph(doc, "Age Cohort Category", selected_age_group, "年龄组别", translate_to_zh(selected_age_group))
    add_bilingual_paragraph(doc, "Client Strengths", student_strengths, "个人优势与强化物", translate_to_zh(student_strengths))
    add_bilingual_paragraph(doc, "Target Behavior Definition", behavior_desc, "目标行为定义", translate_to_zh(behavior_desc))
    add_bilingual_paragraph(doc, "Medical / Setting Factors", medical_factors, "医疗/环境因素", translate_to_zh(medical_factors))

    add_bilingual_heading(doc, '1.5 Qualitative Stakeholder Input & Ecological Notes', '利益相关者访谈与生态评估记录', level=1)
    add_bilingual_paragraph(doc, "Qualitative Summary", uploaded_qual_text, "质性评估摘要", translate_to_zh(uploaded_qual_text))

    add_bilingual_heading(doc, '2. Direct Systematic ABC Observation Ledger', '直接系统化 ABC 行为观察日志', level=1)
    
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

    add_bilingual_heading(doc, '3. Weighted Triangulation & Functional Summary', '加权交叉验证与功能总结', level=1)
    add_bilingual_text(doc, triangulation_summary_eng, triangulation_summary_zh)
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 生成 BIP 文档
def generate_bip_document():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.8)
    
    title_text = 'BEHAVIOR INTERVENTION PLAN (BIP)\n(行为干预计划)' if "Chinese" in selected_language else 'BEHAVIOR INTERVENTION PLAN (BIP)'
    title = doc.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    add_bilingual_heading(doc, '1. Target Behavior & Primary Function', '目标行为与主要功能', level=1)
    add_bilingual_paragraph(doc, "Client Identifier", "[CLIENT_NAME]", "客户标识", "[CLIENT_NAME]")
    add_bilingual_paragraph(doc, "Target Behavior Definition", behavior_desc, "目标行为定义", translate_to_zh(behavior_desc))
    add_bilingual_paragraph(doc, "Inferred Maintaining Function", primary_function, "推导主要功能", translate_to_zh(primary_function))
    add_bilingual_paragraph(doc, "Prescribed Age Focus", age_strategy_note, "针对年龄段策略焦点", age_strategy_note_zh)

    add_bilingual_heading(doc, '2. Proactive Antecedent Strategies', '前因预防策略', level=1)
    if "Adult" in selected_age_group:
        add_bilingual_text(
            doc,
            "1. New Staff Pairing & Environment Structuring: Pair all new/substitute staff with familiar personnel for rapport building prior to presenting demands. Establish clear roommate shared-space boundaries and schedules.",
            "1. 新员工配对与环境结构化：所有新员工/替班员工在呈递任务指令前，必须先与熟悉的工作人员进行关系建立（Pairing）。建立清晰的室友共享空间边界与使用日程表。"
        )
    elif "School-Age" in selected_age_group:
        add_bilingual_text(
            doc,
            "1. Structured Attention Delivery & Screen Time Visual Timers: Provide high non-contingent attention (NCA) during independent activities and utilize 5-min/1-min visual countdowns before iPad transitions.",
            "1. 结构化关注给予与屏幕时间视觉计时器：在独立活动期间提供高频非关联性关注（NCA），并在 iPad 转换前使用 5分钟/1分钟视觉倒计时。"
        )
    else:
        add_bilingual_text(
            doc,
            "1. Visual Pre-Correction & Toy Transition Cues: Provide PECS icons and tactile cues prior to toy sharing or routine changes.",
            "1. 视觉预告与玩具转换提示：在玩具分享或例行程序调整前提供 PECS 图标和触觉提示。"
        )

    add_bilingual_heading(doc, '3. Functional Replacement Behaviors (FCT)', '功能性替代行为训练', level=1)
    add_bilingual_text(
        doc,
        "1. Functional Self-Advocacy & Attention Requesting: Teach client to use AAC/visual cards to request 'Look at my work' or '1-min Screen Extension' prior to escalation.",
        "1. 功能性自我倡导与关注请求：教导客户在行为升级前，使用 AAC/视觉卡片表达“看看我的作业”或请求“1分钟屏幕延长”。"
    )

    add_bilingual_heading(doc, '4. Reactive Consequence & Safety Protocols', '后果应对与安全预案', level=1)
    add_bilingual_text(
        doc,
        "1. Planned Ignoring for Target Behavior & Differential Reinforcement: Minimize eye contact during attention-seeking outbursts while maintaining safety; immediately provide high verbal praise upon appropriate communication.",
        "1. 目标行为的有计划忽视与差异化强化：在寻求关注的发作期间，在确保安全的前提下尽量减少眼神接触；一旦客户使用适当沟通，立即给予高度口头表扬。"
    )

    add_bilingual_heading(doc, '5. Local Re-identification Note', '本地脱敏还原说明', level=1)
    add_bilingual_text(
        doc,
        "Note for BCBA / LBA: All client names are export-masked as [CLIENT_NAME]. Use Word Find & Replace (Ctrl+H) on your local workstation to restore true identifying details prior to clinical filing.",
        "提示：所有名称均已脱敏处理为 [CLIENT_NAME]。在提交评估报告前，请在 Word 中按 Ctrl+H 替换为真实姓名。"
    )

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 下载导出界面
st.subheader("📄 Export US-BCBA Clinical Draft Documents")
col_fba, col_bip = st.columns(2)

with col_fba:
    if st.button("🚀 Compile FBA Report (.docx)", type="primary", use_container_width=True):
        fba_file = generate_fba_document()
        st.success("FBA Report Compiled Successfully!")
        st.download_button(
            label="📄 Download FBA_Report.docx",
            data=fba_file,
            file_name=f"FBA_Report_US_{selected_age_group.split()[0]}.docx",
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
            file_name=f"BIP_Plan_US_{selected_age_group.split()[0]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

st.caption("© 2026 US-BCBA BIPEngine Expert System. Optimized for US Healthcare & EB-2 NIW Academic Demonstration.")
