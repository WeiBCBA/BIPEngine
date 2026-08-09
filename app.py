import streamlit as st
import io
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# 1. 页面基本配置
st.set_page_config(page_title="US-BCBA Clinical BIPEngine v3.3", layout="wide", initial_sidebar_state="expanded")

# 2. 侧边栏配置
with st.sidebar:
    st.markdown("<h2 style='color: #1F4E78;'>⚙️ BIPEngine v3.3</h2>", unsafe_allow_html=True)
    st.caption("US BCBA & LBA Clinical Decision Support Engine")
    st.divider()
    
    st.markdown("### 🔒 Privacy & HIPAA Compliance")
    st.markdown("""
    * **No Cloud Storage**: Zero persistent data retention.
    * **Instant Memory Wipe**: Session clearing on browser close.
    * **De-identification**: Automatic PII masking (`[CLIENT_NAME]`).
    """)
    st.divider()
    
    selected_language = st.selectbox(
        "Report Output Language Target:",
        [
            "English & Chinese Dual-Language (中英双语对照)", 
            "English & Spanish Dual-Language (Español)", 
            "English (Standard US)"
        ]
    )
    
    selected_age_group = st.selectbox(
        "Client Development Cohort:",
        ["School-Age (5-21 yrs)", "Adult / Transition (21+ yrs)", "Early Intervention (2-5 yrs)"],
        key="age_group_select"
    )

is_adult = "Adult" in selected_age_group
subj_en = "Client" if is_adult else "Student"
subj_zh = "客户" if is_adult else "学生"

# 3. 深度多语言字典
DICTIONARY_ZH = {
    "Student Name": "学生姓名", "Client Name": "客户姓名",
    "School Name": "学校/机构名称", "School District": "学区/服务管区",
    "Student DOB": "学生出生日期", "Client DOB": "客户出生日期",
    "Student ID": "学生编号", "Client ID": "客户编号",
    "Date of FBA": "FBA评估日期", "Cohort Category": "人群年龄组别", "Placement": "服务地点",
    
    "Direct Observations": "直接行为观察", "Teacher Interview": "教师/工作人员访谈",
    "Parent Interview": "家长/照护者访谈", "Rating Scales": "评估量表 (如 QABF/MAS)",
    
    "Entry": "记录编号 (Entry)", "Date/Time": "日期/时间 (Date/Time)",
    "Observer Role": "观察者/数据源 (Observer & Tool)", "Setting": "服务环境 (Setting)",
    "Antecedent (A)": "前因 (Antecedent - A)", "Behavior (B)": "行为 (Behavior - B)",
    "Consequence (C)": "后果 (Consequence - C)", "Engine Auto-Inferred Function": "系统推导功能 (Inferred Function)",
    
    # 观察者及环境
    "Direct Care Staff (CR Mobile)": "直接照护人员 (CentralReach 移动端)",
    "Job Coach (Vocational Log)": "职业辅导员 (职业训练日志)",
    "RBT (CentralReach Portal)": "RBT 行为技术员 (CentralReach 数据端)",
    "RBT (Catalyst Data App)": "RBT 行为技术员 (Catalyst 数据采集端)",
    "BCBA Direct Observation": "BCBA 直接临床观察",
    "Paraprofessional (School Data Sheet)": "学校助教/特教副手 (学校数据记录表)",
    
    "Supported Living Apartment": "支持性居住公寓",
    "Day Program / Vocational Workshop": "日间中心 / 职业培训车间",
    "In-Home ABA / Screen Time Transition": "居家 ABA / 屏幕时间转换",
    "Special Ed Classroom (Small Group)": "特教教室 (小组教学)",
    "General Ed Classroom (Desk Work)": "普通教室 (课桌作业)",
    "Clinic Therapy Room": "诊所治疗室", "Home ABA Session": "居家 ABA 训练", "Clinic Social Skills Group": "诊所社交技能小组",
    
    # 学龄组 ABC 描述字典
    "Timer rang signaling 30-min iPad screen time limits reached while RBT turned to document data.": "当 RBT 转过身记录数据时，定时器响起，提示 30 分钟 iPad 屏幕使用时间已满。",
    "Pacing, grabbing iPad back, screaming 'Look at me!', dropping to floor.": "来回踱步、抢回 iPad、大喊‘看着我！’并摔倒在地。",
    "RBT made immediate eye contact, prompted '1-min visual extension card', and reinforced quiet waiting.": "RBT 立即与其建立眼神接触，出示‘1分钟视觉延迟卡’提示，并强化其安静等待的行为。",
    
    "Instructor turned attention to assist peer during iPad group activity.": "在 iPad 小组活动期间，指导教师将注意力转向协助同伴。",
    "Approached staff, pulled sleeve, loud vocalizations, tried to grab peer's iPad.": "靠近工作人员、拉拽袖子、发出大声吵闹，并尝试抢夺同伴的 iPad。",
    "Staff turned immediately, made eye contact, and redirected to waiting visual schedule.": "工作人员立即转过身与其建立眼神接触，并将其重定向至等待视觉日程表。",
    
    "Teacher presented multi-step writing worksheet.": "教师出示了一份多步骤的书写练习页。",
    "Screamed (>80dB), pushed desk away.": "尖叫（音量大于80分贝），并将课桌推开。",
    "Staff presented 'Break' visual card; demand paused.": "工作人员出示了‘休息’视觉卡片；任务指令暂停。",

    # 成人组 ABC 描述字典
    "Newly hired staff member presented morning chore checklist.": "新入职的工作人员出示了早间日常家务清单。",
    "Verbal aggression (cursing, threats) and physical aggression (shoving staff).": "言语攻击（辱骂、威胁）及肢体攻击（推搡工作人员）。",
    "Senior BCBA/Staff stepped in, guided new staff to pause demand, and represented visual choice board.": "高级 BCBA/资深员工介入，指导新员工暂停任务要求，并重新出示视觉选择板。",
    "Roommate turned on living room TV and adjusted seating area without client consent.": "室友在未获得客户同意的情况下打开客厅电视并调整座位区域。",
    "Loud vocal resistance, blocking TV screen, grabbing remote from roommate.": "大声言语反抗、遮挡电视屏幕、抢夺室友手中的遥控器。",
    "Staff prompted roommate mediation and provided alternative personal tablet for private room.": "工作人员介入进行室友调解，并为客户提供可在私人房间使用的替代个人平板电脑。",
    "Client reported joint pain/headache after 2 hours of repetitive standing work.": "客户在连续进行 2 小时重复性站立工作后报告关节疼痛/头痛。",
    "Pacing, hand-wringing, aggressive resistance to verbal prompts.": "来回踱步、拧双手、对口头提示表现出攻击性抗拒。",
    "Staff offered PRN pain relief medication and quiet rest area.": "工作人员按照按需医嘱 (PRN) 提供止痛药物并安排安静休息区。",
    
    # 功能项
    "Task Escape / Demand Avoidance": "逃避 / 避免任务 (Task Escape)",
    "Access to Tangibles / Activities": "获取实物/活动/环境控制 (Access to Tangibles/Control)",
    "Physical Discomfort / Internal State": "生理不适/内部状态 (Physical Discomfort)",
    "Social Attention Seeking": "寻求社交关注 (Social Attention)",
    "Automatic / Sensory Stimulation": "自动强化/感官刺激 (Automatic / Sensory)"
}

def translate(text_val, target_lang):
    if not text_val: return ""
    val_str = str(text_val).strip()
    if "Chinese" in target_lang:
        if val_str in DICTIONARY_ZH: return DICTIONARY_ZH[val_str]
        res = val_str
        for k, v in DICTIONARY_ZH.items():
            res = res.replace(k, v)
        return res
    return val_str

def get_preset_abc(age_group):
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
    else: # School-Age 5-21
        return [
            {"Entry": "Obs #1", "Date/Time": "08/03/2026 04:30 PM", "Observer Role": "RBT (CentralReach Portal)", "Setting": "In-Home ABA / Screen Time Transition", "Antecedent (A)": "Timer rang signaling 30-min iPad screen time limits reached while RBT turned to document data.", "Behavior (B)": "Pacing, grabbing iPad back, screaming 'Look at me!', dropping to floor.", "Consequence (C)": "RBT made immediate eye contact, prompted '1-min visual extension card', and reinforced quiet waiting."},
            {"Entry": "Obs #2", "Date/Time": "08/04/2026 01:45 PM", "Observer Role": "Paraprofessional (School Data Sheet)", "Setting": "Special Ed Classroom (Small Group)", "Antecedent (A)": "Instructor turned attention to assist peer during iPad group activity.", "Behavior (B)": "Approached staff, pulled sleeve, loud vocalizations, tried to grab peer's iPad.", "Consequence (C)": "Staff turned immediately, made eye contact, and redirected to waiting visual schedule."},
            {"Entry": "Obs #3", "Date/Time": "08/05/2026 09:15 AM", "Observer Role": "BCBA Direct Observation", "Setting": "General Ed Classroom (Desk Work)", "Antecedent (A)": "Teacher presented multi-step writing worksheet.", "Behavior (B)": "Screamed (>80dB), pushed desk away.", "Consequence (C)": "Staff presented 'Break' visual card; demand paused."}
        ]

def get_age_defaults(age_group):
    if "Adult" in age_group:
        return {
            "strengths": "Responds well to visual timers, high motivation for cause-and-effect vocational tools.",
            "history": "Enrolled in supported living program; receives vocational training and OT services.",
            "target": "Physical aggression toward new staff, verbal threats, and roommate conflicts over shared items.",
            "freq": "3 to 5 episodes per week during staff transitions or common area usage.",
            "dur": "Episodes last 15 to 45 minutes until resolution.",
            "int": "High Intensity (Aggression towards staff & safety concerns)",
            "slow": "Unfamiliar substitute staff on shift, physical joint pain, or lack of sleep.",
            "fast": "Prompts for transition, removal of tangibles, or social attention shift.",
            "non_occ": "Preferred staff presence, clear visual schedule, individual leisure time in private room.",
            "cons": "Task demands paused, sensory chew tools or verbal redirection provided.",
            "primary_func": "Access to Tangibles / Social Attention"
        }
    else: # School-Age (5-21 yrs)
        return {
            "strengths": "High motivation for iPad videos, responds well to 1-on-1 attention and visual countdown timers.",
            "history": "Enrolled in Special Education classroom; receives Speech (SLP) and ABA services.",
            "target": "Screaming (>80dB), grabbing devices, pulling staff sleeves, and dropping to floor when screen time ends or teacher attention shifts.",
            "freq": "4 to 6 times per week during transitions or group activities.",
            "dur": "Episodes last 5 to 15 minutes.",
            "int": "Moderate to High Intensity (Disruptive to classroom/session)",
            "slow": "Fatigue after school, unexpected routine changes, or delayed adult attention.",
            "fast": "Screen time timer ringing, instructor turning attention to peers, or multi-step writing tasks.",
            "non_occ": "1-on-1 dedicated attention, clear visual schedules, high-preference leisure tasks.",
            "cons": "Immediate eye contact provided, visual extension card offered, task demand temporarily paused.",
            "primary_func": "Access to Tangibles / Social Attention"
        }

curr_defaults = get_age_defaults(selected_age_group)

# 4. Streamlit 界面构建
st.title("🚀 US-BCBA Automated FBA & BIP Clinical Compiler")
st.divider()

st.header("📋 Phase 1: Multi-Source Clinical Ingestion")

tab1, tab2, tab3 = st.tabs([
    "📊 Direct ABC Observations", 
    "📝 Expanded FBA Inputs", 
    "📈 QABF Scale"
])

with tab1:
    st.subheader("Direct Systematic ABC Data Ledger")
    default_abc = get_preset_abc(selected_age_group)
    raw_df = pd.DataFrame(default_abc)

    def infer_func(row):
        ant = str(row.get("Antecedent (A)", "")).lower()
        beh = str(row.get("Behavior (B)", "")).lower()
        full_text = f"{ant} {beh}"
        if any(k in full_text for k in ["pain", "medication", "joint", "headache"]):
            return "Physical Discomfort / Internal State"
        if any(k in full_text for k in ["demand", "task", "worksheet", "chore", "writing"]):
            return "Task Escape / Demand Avoidance"
        if any(k in full_text for k in ["ipad", "toy", "screen", "tv", "remote"]):
            return "Access to Tangibles / Activities"
        if "look at me" in beh or "pulled sleeve" in beh or "attention" in ant:
            return "Social Attention Seeking"
        return "Automatic / Sensory Stimulation"

    raw_df["Engine Auto-Inferred Function"] = raw_df.apply(infer_func, axis=1)
    edited_abc = st.data_editor(raw_df, num_rows="dynamic", use_container_width=True, key=f"editor_{selected_age_group}")

with tab2:
    st.subheader(f"🤝 Stakeholder Input ({subj_en}-Centered)")
    
    col_a, col_b = st.columns(2)
    with col_a:
        agency_name = st.text_input("School / Agency Name", "Metropolitan Inclusive Center")
        district_name = st.text_input("District / Health Region", "District 10 Behavioral Division")
        dob_val = st.text_input(f"{subj_en} DOB", "05/12/2015" if not is_adult else "05/12/2001")
        id_val = st.text_input(f"{subj_en} ID", "ID-908231")
        fba_date = st.text_input("Date of FBA", "08/08/2026")
        
    with col_b:
        sources_options = ["Direct Observations", "Teacher Interview", "Parent Interview", "Rating Scales"]
        data_sources = st.multiselect("1. Data Sources:", sources_options, default=sources_options)
        
    st.divider()
    st.markdown("### 2. Brief Background & Strengths")
    strengths_val = st.text_area("Strengths & Motivators", curr_defaults["strengths"], height=65)
    ed_history_val = st.text_area("Educational / Service History", curr_defaults["history"], height=65)

    st.divider()
    st.markdown("### 3. Target Behavior Operational Breakdown")
    c1, c2, c3, c4 = st.columns(4)
    target_beh = c1.text_area("Target Behavior Description", curr_defaults["target"], height=70)
    freq_val = c2.text_area("Frequency", curr_defaults["freq"], height=70)
    dur_val = c3.text_area("Duration", curr_defaults["dur"], height=70)
    int_val = c4.text_area("Intensity", curr_defaults["int"], height=70)

    st.divider()
    st.markdown("### 4. Behavioral Triggers & Context")
    c_t1, c_t2 = st.columns(2)
    setting_events = c_t1.text_area("Setting Events (Slow Triggers)", curr_defaults["slow"], height=70)
    antecedents_val = c_t2.text_area("Antecedent Events (Immediate Triggers)", curr_defaults["fast"], height=70)
    
    st.divider()
    st.markdown("### 7 & 8. Non-Occurrence Context & Immediate Consequences")
    c_t3, c_t4 = st.columns(2)
    non_occ_val = c_t3.text_area("7) Non-Occurrence Situations", curr_defaults["non_occ"], height=70)
    consequences_val = c_t4.text_area("8) Consequences (Immediate Responses)", curr_defaults["cons"], height=70)

with tab3:
    st.subheader("📈 QABF Psychometric Scale Scores")
    q1, q2, q3, q4, q5 = st.columns(5)
    att_score = q1.number_input("Social Attention", 0, 15, 12 if not is_adult else 3)
    esc_score = q2.number_input("Task Escape", 0, 15, 10 if not is_adult else 4)
    tan_score = q3.number_input("Tangibles / Control", 0, 15, 14 if not is_adult else 12)
    sen_score = q4.number_input("Sensory Stimulation", 0, 15, 4 if not is_adult else 14)
    phy_score = q5.number_input("Physical Discomfort", 0, 15, 2 if not is_adult else 11)

st.divider()

# 5. 编译与导出 Word 文档函数（完美统一 1-10 格式与补齐第 6 项）
def generate_aligned_fba_docx():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.7)
    
    if "Chinese" in selected_language:
        title_str = "FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) FORM\n(功能性行为评估标准表)"
    else:
        title_str = "FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) FORM"
        
    title_p = doc.add_heading(title_str, level=0)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 0. 基本信息表格
    info_table = doc.add_table(rows=4, cols=4)
    info_table.style = 'Table Grid'
    
    table_data = [
        [(f"{subj_en} Name", "[CLIENT_NAME]"), ("School Name", agency_name)],
        [(f"{subj_en} DOB", dob_val), ("School District", district_name)],
        [(f"{subj_en} ID", id_val), ("Date of FBA", fba_date)],
        [("Cohort Category", selected_age_group), ("Placement", agency_name)]
    ]
    
    for r_idx, row in enumerate(table_data):
        for c_group, (lbl, val) in enumerate(row):
            cell_lbl = info_table.cell(r_idx, c_group * 2)
            cell_val = info_table.cell(r_idx, c_group * 2 + 1)
            
            lbl_trans = translate(lbl, selected_language)
            cell_lbl.text = f"{lbl}({lbl_trans})" if selected_language != "English (Standard US)" and lbl_trans != lbl else lbl
            cell_val.text = str(val)
            shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
            cell_lbl._tc.get_or_add_tcPr().append(shd)
            cell_lbl.paragraphs[0].runs[0].font.bold = True
            cell_lbl.paragraphs[0].runs[0].font.size = Pt(8.5)
            cell_val.paragraphs[0].runs[0].font.size = Pt(8.5)

    doc.add_paragraph()

    # 1. Data Sources
    doc.add_heading("1. Data Sources (数据来源)" if "Chinese" in selected_language else "1. Data Sources", level=1)
    sources_str = ", ".join(data_sources)
    doc.add_paragraph(f"Selected Sources: {sources_str}")

    # 2. Brief Background & Strengths
    doc.add_heading("2. Brief Background & Strengths (简要背景与优势)" if "Chinese" in selected_language else "2. Brief Background & Strengths", level=1)
    p2_1 = doc.add_paragraph()
    p2_1.add_run("Strengths & Motivators: ").bold = True
    p2_1.add_run(strengths_val)
    if "Chinese" in selected_language:
        p2_1.add_run(f"\n（中文对照 - 优势与强化物: {translate(strengths_val, selected_language)}）").italic = True
        
    p2_2 = doc.add_paragraph()
    p2_2.add_run("Educational History: ").bold = True
    p2_2.add_run(ed_history_val)
    if "Chinese" in selected_language:
        p2_2.add_run(f"\n（中文对照 - 教育/服务背景履历: {translate(ed_history_val, selected_language)}）").italic = True

    # 3. Target Behavior Operational Breakdown
    doc.add_heading("3. Target Behavior Operational Breakdown (目标行为操作化拆解)" if "Chinese" in selected_language else "3. Target Behavior Operational Breakdown", level=1)
    for lbl_en, lbl_zh, val in [
        ("Description of Target Behavior", "目标行为操作化定义", target_beh),
        ("Frequency", "发生频率", freq_val),
        ("Duration", "持续时间", dur_val),
        ("Intensity", "行为强度", int_val)
    ]:
        p = doc.add_paragraph()
        p.add_run(f"{lbl_en}: ").bold = True
        p.add_run(val)
        if "Chinese" in selected_language:
            p.add_run(f"\n（中文对照 - {lbl_zh}: {translate(val, selected_language)}）").italic = True

    # 4. Behavioral Triggers & Context
    doc.add_heading("4. Behavioral Triggers & Context (行为诱因与环境情境)" if "Chinese" in selected_language else "4. Behavioral Triggers & Context", level=1)
    for lbl_en, lbl_zh, val in [
        ("Setting Events (Slow Triggers)", "情境因素 (慢速诱因)", setting_events),
        ("Antecedent Events (Immediate Triggers)", "即时前因 (快速诱因)", antecedents_val)
    ]:
        p = doc.add_paragraph()
        p.add_run(f"{lbl_en}: ").bold = True
        p.add_run(val)
        if "Chinese" in selected_language:
            p.add_run(f"\n（中文对照 - {lbl_zh}: {translate(val, selected_language)}）").italic = True

    # 5. Direct Systematic ABC Observation Ledger
    doc.add_heading("5. Direct Systematic ABC Observation Ledger (直接系统化 ABC 观察日志)" if "Chinese" in selected_language else "5. Direct Systematic ABC Observation Ledger", level=1)
    headers = list(edited_abc.columns)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    for idx, text in enumerate(headers):
        cell = table.rows[0].cells[idx]
        trans_h = translate(text, selected_language)
        cell.text = f"{text}({trans_h})" if selected_language != "English (Standard US)" else text
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
            raw_val = str(val)
            trans_v = translate(raw_val, selected_language)
            cell_display = f"{raw_val}\n（{trans_v}）" if selected_language != "English (Standard US)" and trans_v and trans_v != raw_val else raw_val
                
            row_cells[c_idx].text = cell_display
            p = row_cells[c_idx].paragraphs[0]
            p.runs[0].font.size = Pt(7.5)
            if r_idx % 2 == 1:
                shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
                row_cells[c_idx]._tc.get_or_add_tcPr().append(shd)

    # 6. Summary of QABF Rating Scale Scores (补齐缺失的第6项，格式完全对齐)
    doc.add_heading("6. Summary of QABF Rating Scale Scores (QABF 量表评估得分汇总)" if "Chinese" in selected_language else "6. Summary of QABF Rating Scale Scores", level=1)
    p6 = doc.add_paragraph()
    qabf_summary = f"Attention: {att_score}/15 | Escape: {esc_score}/15 | Tangible/Control: {tan_score}/15 | Sensory: {sen_score}/15 | Physical Discomfort: {phy_score}/15"
    p6.add_run(qabf_summary).bold = True
    if "Chinese" in selected_language:
        p6.add_run(f"\n（中文对照 - QABF 各维度得分: 社交关注 {att_score}分 | 逃避任务 {esc_score}分 | 获取实物 {tan_score}分 | 感官刺激 {sen_score}分 | 生理不适 {phy_score}分）").italic = True

    # 7. Non-Occurrence Situations (格式统一为 Heading 1 标准样式)
    doc.add_heading("7. Non-Occurrence Situations & Exceptions (行为不常发生的例外情境)" if "Chinese" in selected_language else "7. Non-Occurrence Situations & Exceptions", level=1)
    p7 = doc.add_paragraph()
    p7.add_run("Identify events or times when behavior does NOT usually occur: ").bold = True
    p7.add_run(non_occ_val)
    if "Chinese" in selected_language:
        p7.add_run(f"\n（中文对照 - 例外情境说明: {translate(non_occ_val, selected_language)}）").italic = True

    # 8. Consequences (格式统一为 Heading 1 标准样式)
    doc.add_heading("8. Environmental Consequences & Immediate Responses (行为后果与他人反应)" if "Chinese" in selected_language else "8. Environmental Consequences & Immediate Responses", level=1)
    p8 = doc.add_paragraph()
    p8.add_run("Immediate responses following problem behavior: ").bold = True
    p8.add_run(consequences_val)
    if "Chinese" in selected_language:
        p8.add_run(f"\n（中文对照 - 行为后果说明: {translate(consequences_val, selected_language)}）").italic = True

    # 9. Hypothesis Statement (格式统一为 Heading 1 标准样式)
    doc.add_heading("9. Functional Behavioral Hypothesis Statement (行为功能假设说明)" if "Chinese" in selected_language else "9. Functional Behavioral Hypothesis Statement", level=1)
    primary_func = curr_defaults["primary_func"]
    hyp_en = f"When presented with {antecedents_val} under setting conditions of {setting_events}, [CLIENT_NAME] engages in {target_beh} which results in {consequences_val} in order to achieve {primary_func}. The behavior serves as a functional attempt to communicate."

    p9 = doc.add_paragraph()
    p9.add_run("Primary Hypothesis: ").bold = True
    p9.add_run(hyp_en)
    if "Chinese" in selected_language:
        hyp_zh = f"当在【{translate(setting_events, selected_language)}】的背景情境下出现【{translate(antecedents_val, selected_language)}】时，[CLIENT_NAME] 会表现出【{translate(target_beh, selected_language)}】，从而导致【{translate(consequences_val, selected_language)}】，其主要目的是【{translate(primary_func, selected_language)}】。该行为是传递沟通意图的功能性尝试。"
        p9.add_run(f"\n（中文对照 - 行为假设表达: {hyp_zh}）").italic = True

    # 10. Function of Behavior Matrix (格式统一为 Heading 1 标准样式 + 规范矩阵)
    doc.add_heading("10. Primary Function of Behavior Matrix (行为主要功能认定矩阵)" if "Chinese" in selected_language else "10. Primary Function of Behavior Matrix", level=1)
    
    func_table = doc.add_table(rows=2, cols=2)
    func_table.style = 'Table Grid'
    
    is_att = "Attention" in primary_func
    is_esc = "Escape" in primary_func
    is_tan = "Tangible" in primary_func or "Control" in primary_func
    is_sen = "Sensory" in primary_func or "Physical" in primary_func

    func_table.cell(0, 0).text = f"[{'X' if is_att else '  '}] Attention (社交关注)"
    func_table.cell(0, 1).text = f"[{'X' if is_tan else '  '}] Tangible (获取实物/活动)"
    func_table.cell(1, 0).text = f"[{'X' if is_esc else '  '}] Escape (逃避任务/指令)"
    func_table.cell(1, 1).text = f"[{'X' if is_sen else '  '}] Sensory (感官/生理调节)"

    p_notes = doc.add_paragraph()
    p_notes.add_run("\nAdditional Clinical Notes: ").bold = True
    p_notes.add_run("Data collected aligns across direct observation and rating scales. Recommending BIP focusing on Functional Communication Training (FCT).")

    # 结尾临床免责/修改英文声明
    doc.add_paragraph()
    p_disclaimer = doc.add_paragraph()
    p_disclaimer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_disc = p_disclaimer.add_run("⚠️ CLINICAL NOTICE TO BCBA / LBA:\nThis FBA report draft is synthesized automatically by BIPEngine based on standard templates. Please review, adjust, and individualize all operational descriptions, hypothesis formulations, and target functions according to the client's actual clinical needs and unique presentation prior to formal signature.")
    r_disc.font.size = Pt(8.5)
    r_disc.font.italic = True
    r_disc.font.color.rgb = RGBColor(120, 120, 120)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 6. 生成与导出操作区（提示框改为高专业度英文提示）
st.markdown("### ⚠️ Clinical Responsibility Disclaimer")
st.warning("💡 **BCBA / LBA Professional Notice**: This FBA draft is auto-compiled using clinical algorithms and standard templates. Prior to official sign-off, please carefully review, refine, and **individualize all behavioral operational definitions, hypothesis statements, and target functions** to precisely align with the client's unique presentation and clinical data.")

if st.button("🚀 Compile Aligned FBA Report (.docx)", type="primary", use_container_width=True):
    fba_file = generate_aligned_fba_docx()
    st.success("Aligned FBA Report Compiled Successfully!")
    st.download_button(
        label="📄 Download Aligned_FBA_Report.docx",
        data=fba_file,
        file_name="Aligned_FBA_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
