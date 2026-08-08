import streamlit as st
import io
import re
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="US-BCBA Clinical BIPEngine v2.6", layout="wide", initial_sidebar_state="expanded")

# 侧边栏配置
with st.sidebar:
    st.markdown("<h2 style='color: #1F4E78;'>⚙️ BIPEngine v2.6 (US Standard)</h2>", unsafe_allow_html=True)
    st.caption("US BCBA & LBA Clinical Decision Support Engine for FBA/BIP")
    st.divider()
    
    selected_language = st.selectbox(
        "Report Output Language Target:",
        ["English & Chinese Dual-Language (中英双语对照)", "English (Standard US)"]
    )
    
    selected_age_group = st.selectbox(
        "Client Development Cohort:",
        ["School-Age (5-21 yrs)", "Early Intervention (2-5 yrs)", "Adult / Transition (21+ yrs)"],
        key="age_group_select"
    )

st.title("🚀 US-BCBA Automated FBA & BIP Clinical Compiler")
st.markdown("##### Advanced Triangulation Engine adhering to BACB & US Healthcare Standards")
st.divider()

# 动态决定主语称呼（Student vs Client）
is_adult = "Adult" in selected_age_group
subj_en = "Client" if is_adult else "Student"
subj_zh = "客户" if is_adult else "学生"

# 全面补全中英双语字典 (含 ABC 完整句子与所有表格字段)
DICTIONARY_ZH = {
    # 动态表头与字段
    "Student Name": "学生姓名",
    "Client Name": "客户姓名",
    "School Name": "学校/机构名称",
    "School District": "学区/服务管区",
    "Student DOB": "学生出生日期",
    "Client DOB": "客户出生日期",
    "Student ID": "学生编号",
    "Client ID": "客户编号",
    "Date of FBA": "FBA评估日期",
    "Cohort Category": "人群年龄组别",
    "Placement": "安置环境/服务地点",
    
    # 数据源
    "Direct Observations": "直接行为观察",
    "Student Interview": "学生访谈",
    "Client Interview": "客户访谈",
    "Teacher Interview": "教师/工作人员访谈",
    "Parent Interview": "家长/照护者访谈",
    "Rating Scales": "评估量表 (如 QABF/MAS)",
    
    # 维度字段
    "Frequency": "发生频率",
    "Duration": "持续时间",
    "Intensity": "行为强度",
    "Setting Events": "情境因素/慢速诱因",
    "Antecedent Events": "即时前因/快速诱因",
    "Non-occurrence Situations": "行为不常发生的例外情境",
    "Consequences": "行为后果",
    "Hypothesis Statement": "行为假设说明",
    "Communicative Intent": "沟通意图",
    "Brief Background": "简要背景履历",
    "Strengths": "优势与强化物",
    
    # ABC 观察点环境与描述（长句映射）
    "Supported Living Apartment": "支持性居住公寓",
    "Day Program / Vocational Workshop": "日间中心 / 职业培训车间",
    "Clinic Therapy Room": "诊所治疗室",
    "Home ABA Session": "居家 ABA 训练",
    "Clinic Social Skills Group": "诊所社交技能小组",
    "In-Home ABA / Screen Time Transition": "居家 ABA / 屏幕时间转换",
    "Special Ed Classroom (Small Group)": "特教教室 (小组教学)",
    "General Ed Classroom (Desk Work)": "普通教室 (课桌作业)",
    
    # ABC Antecedent / Behavior / Consequence 长句精确翻译
    "Newly hired staff member presented morning chore checklist.": "新入职的工作人员出示了早间日常家务清单。",
    "Verbal aggression (cursing, threats) and physical aggression (shoving staff).": "言语攻击（辱骂、威胁）及肢体攻击（推搡工作人员）。",
    "Senior BCBA/Staff stepped in, guided new staff to pause demand, and represented visual choice board.": "高级 BCBA/资深员工介入，指导新员工暂停任务要求，并重新出示视觉选择板。",
    "Roommate turned on living room TV and adjusted seating area without client consent.": "室友在未获得客户同意的情况下打开客厅电视并调整座位区域。",
    "Loud vocal resistance, blocking TV screen, grabbing remote from roommate.": "大声言语反抗、遮挡电视屏幕、抢夺室友手中的遥控器。",
    "Staff prompted roommate mediation and provided alternative personal tablet for private room.": "工作人员介入进行室友调解，并为客户提供可在私人房间使用的替代个人平板电脑。",
    "Client reported joint pain/headache after 2 hours of repetitive standing work.": "客户在连续进行 2 小时重复性站立工作后报告关节疼痛/头痛。",
    "Pacing, hand-wringing, aggressive resistance to verbal prompts.": "来回踱步、拧双手、对口头提示表现出攻击性抗拒。",
    "Staff offered PRN pain relief medication and quiet rest area.": "工作人员按照按需医嘱 (PRN) 提供止痛药物并安排安静休息区。",
    
    "RBT requested sharing toy truck during naturalistic play.": "RBT 在自然情境游戏期间要求分享玩具卡车。",
    "Screamed, bit own arm, threw toy.": "尖叫、咬自己的手臂、扔玩具。",
    "RBT paused demand, offered sensory chew tool.": "RBT 暂停任务指令，提供感官咀咀嚼咬乐工具。",
    "Therapist transitioned from bubble play to discrete trial teaching (DTT).": "治疗师从吹泡泡游戏转换至分解尝试教导 (DTT)。",
    "Dropped to floor, crying, head banging on carpet.": "瘫倒在地、哭泣、在地毯上用头撞地。",
    "Therapist paused demand, presented PECS break icon.": "治疗师暂停任务指令，出示 PECS 休息图标。",
    "RBT called for group cleanup time.": "RBT 呼叫进行小组收拾玩具时间。",
    "Ran toward clinic exit door (elopement).": "跑向诊所出口大门（离位/逃跑）。",
    "Staff guided back with visual transition timer.": "工作人员通过视觉过渡计时器将其引导返回。",
    
    "Timer rang signaling 30-min iPad screen time limits reached while RBT turned to document data.": "当 RBT 转头记录数据时，定时器响起提示 30 分钟 iPad 屏幕时间已结束。",
    "Pacing, grabbing iPad back, screaming 'Look at me!', dropping to floor.": "来回踱步、抢回 iPad、尖叫“看我！”并瘫倒在地。",
    "RBT made immediate eye contact, prompted '1-min visual extension card', and reinforced quiet waiting.": "RBT 立即建立眼神接触，出示“1分钟延时卡”，并在其安静等待后给予强化。",
    "Instructor turned attention to assist peer during iPad group activity.": "在 iPad 小组活动期间，指导员将注意力转向协助同伴。",
    "Approached staff, pulled sleeve, loud vocalizations, tried to grab peer's iPad.": "靠近工作人员、拉扯袖子、大声发出声音，并试图抢夺同伴的 iPad。",
    "Staff turned immediately, made eye contact, and redirected to waiting visual schedule.": "工作人员立即转头建立眼神接触，并将其引导至等待视觉日程表。",
    "Teacher presented multi-step writing worksheet.": "教师出示多步骤书写工作单。",
    "Screamed (>80dB), pushed desk away.": "尖叫（音量>80分贝），推开课桌。",
    "Staff presented 'Break' visual card; demand paused.": "工作人员出示“休息”视觉卡片；任务暂停。",
    
    # 功能描述
    "Task Escape / Demand Avoidance": "逃避 / 避免任务 (Task Escape)",
    "Access to Tangibles / Activities": "获取实物/活动/环境控制 (Access to Tangibles/Control)",
    "Social Attention Seeking": "寻求社交关注 (Social Attention)",
    "Automatic / Sensory Stimulation": "自动强化/感官刺激 (Automatic / Sensory)",
    "Physical Discomfort / Internal State": "生理不适/内部状态 (Physical Discomfort)"
}

def translate_text(text_val):
    if not text_val:
        return ""
    val_str = str(text_val).strip()
    if val_str in DICTIONARY_ZH:
        return DICTIONARY_ZH[val_str]
    # 如果找不到精确匹配，进行长句片段替换
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
    else:
        return [
            {"Entry": "Obs #1", "Date/Time": "08/03/2026 04:30 PM", "Observer Role": "RBT (CentralReach Portal)", "Setting": "In-Home ABA / Screen Time Transition", "Antecedent (A)": "Timer rang signaling 30-min iPad screen time limits reached while RBT turned to document data.", "Behavior (B)": "Pacing, grabbing iPad back, screaming 'Look at me!', dropping to floor.", "Consequence (C)": "RBT made immediate eye contact, prompted '1-min visual extension card', and reinforced quiet waiting."},
            {"Entry": "Obs #2", "Date/Time": "08/04/2026 01:45 PM", "Observer Role": "Paraprofessional (School Data Sheet)", "Setting": "Special Ed Classroom (Small Group)", "Antecedent (A)": "Instructor turned attention to assist peer during iPad group activity.", "Behavior (B)": "Approached staff, pulled sleeve, loud vocalizations, tried to grab peer's iPad.", "Consequence (C)": "Staff turned immediately, made eye contact, and redirected to waiting visual schedule."},
            {"Entry": "Obs #3", "Date/Time": "08/05/2026 09:15 AM", "Observer Role": "BCBA Direct Observation", "Setting": "General Ed Classroom (Desk Work)", "Antecedent (A)": "Teacher presented multi-step writing worksheet.", "Behavior (B)": "Screamed (>80dB), pushed desk away.", "Consequence (C)": "Staff presented 'Break' visual card; demand paused."}
        ]

st.header("📋 Phase 1: Multi-Source Clinical Ingestion")
tab1, tab2 = st.tabs(["📊 Direct ABC Observations", "📝 Expanded FBA Template Inputs"])

with tab1:
    default_abc_data = get_age_specific_abc_presets(selected_age_group)
    raw_df = pd.DataFrame(default_abc_data)

    def infer_function_from_row(row):
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

    raw_df["Engine Auto-Inferred Function"] = raw_df.apply(infer_function_from_row, axis=1)
    edited_abc = st.data_editor(raw_df, num_rows="dynamic", use_container_width=True, key=f"editor_{selected_age_group}")

with tab2:
    st.subheader(f"Expanded FBA Form Fields ({subj_en}-Centered Format)")
    col_a, col_b = st.columns(2)
    with col_a:
        agency_name = st.text_input("School / Agency Name", "Metropolitan Inclusive Center")
        district_name = st.text_input("District / Health Region", "District 10 Behavioral Division")
        dob_val = st.text_input(f"{subj_en} DOB", "05/12/2021")
        id_val = st.text_input(f"{subj_en} ID", "ID-908231")
        fba_date = st.text_input("Date of FBA", "08/08/2026")
        
    with col_b:
        sources_options = [f"{subj_en} Interview", "Direct Observations", "Teacher/Staff Interview", "Parent/Caregiver Interview", "Rating Scales"]
        data_sources = st.multiselect("Data Sources:", sources_options, default=["Direct Observations", "Teacher/Staff Interview", "Parent/Caregiver Interview"])
        background_info = st.text_area("Brief Background", "Receives support services across daily program settings.", height=68)

    st.divider()
    c1, c2, c3 = st.columns(3)
    behavior_desc = c1.text_area("Target Behavior Description", "Verbal threats, shoving staff during transitions.", height=80)
    freq_desc = c2.text_area("Frequency", "3 to 5 episodes per week.", height=80)
    dur_desc = c3.text_area("Duration", "15 to 45 minutes per episode.", height=80)

# 生成完整完美翻译的 Word 文档
def generate_fba_document():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.7)
    
    # 动态标题
    title_text = f'FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) FORM\n({subj_zh}功能性行为评估标准表)' if "Chinese" in selected_language else 'FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) FORM'
    title = doc.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 1. 动态表头网格 (支持 Student vs Client 术语切换 + 精准中文翻译)
    header_table = doc.add_table(rows=4, cols=4)
    header_table.style = 'Table Grid'
    
    info_map = [
        [(f"{subj_en} Name", "[CLIENT_NAME]"), ("School/Agency Name", agency_name)],
        [(f"{subj_en} DOB", dob_val), ("School District/Region", district_name)],
        [(f"{subj_en} ID", id_val), ("Date of FBA", fba_date)],
        [("Cohort Category", selected_age_group), ("Placement Location", agency_name)]
    ]
    
    for r_idx, row_data in enumerate(info_map):
        for c_group, (lbl_en, val) in enumerate(row_data):
            c_lbl = header_table.cell(r_idx, c_group * 2)
            c_val = header_table.cell(r_idx, c_group * 2 + 1)
            
            zh_lbl = translate_text(lbl_en)
            c_lbl.text = f"{lbl_en}\n({zh_lbl})" if "Chinese" in selected_language else lbl_en
            c_val.text = str(val)
            
            shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
            c_lbl._tc.get_or_add_tcPr().append(shd)
            c_lbl.paragraphs[0].runs[0].font.bold = True
            c_lbl.paragraphs[0].runs[0].font.size = Pt(8.5)
            c_val.paragraphs[0].runs[0].font.size = Pt(8.5)

    doc.add_paragraph()

    # 2. ABC 表格导出 (含单元格内容自动翻译)
    doc.add_heading('Direct Systematic ABC Observation Ledger (直接系统化 ABC 观察日志)', level=1)
    headers = list(edited_abc.columns)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    for idx, text in enumerate(headers):
        cell = table.rows[0].cells[idx]
        cell.text = f"{text}\n({translate_text(text)})" if "Chinese" in selected_language else str(text)
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
            zh_val = translate_text(raw_val) # 核心修复：长句与词条全面查表翻译
            
            if "Chinese" in selected_language and zh_val and zh_val != raw_val:
                cell_display = f"{raw_val}\n（{zh_val}）"
            else:
                cell_display = raw_val
                
            row_cells[c_idx].text = cell_display
            p = row_cells[c_idx].paragraphs[0]
            p.runs[0].font.size = Pt(7.5)
            if r_idx % 2 == 1:
                shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
                row_cells[c_idx]._tc.get_or_add_tcPr().append(shd)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

st.divider()
if st.button("🚀 Compile Perfect Dual-Language FBA (.docx)", type="primary", use_container_width=True):
    fba_file = generate_fba_document()
    st.success("FBA Report Compiled with Full Dual-Language Mapping & Dynamic Terminology!")
    st.download_button(
        label="📄 Download Fixed_Bilingual_FBA_Report.docx",
        data=fba_file,
        file_name=f"Fixed_FBA_Report_{selected_age_group.split()[0]}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
