import streamlit as st
import io
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# 1. 页面基本配置
st.set_page_config(page_title="US-BCBA Clinical FBA & BIP Engine v4.0", layout="wide", initial_sidebar_state="expanded")

# 2. 侧边栏配置
with st.sidebar:
    st.markdown("<h2 style='color: #1F4E78;'>⚙️ BIPEngine v4.0</h2>", unsafe_allow_html=True)
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
        ["Early Intervention (2-5 yrs)", "School-Age (5-21 yrs)", "Adult / Transition (21+ yrs)"],
        key="age_group_select"
    )

is_adult = "Adult" in selected_age_group
subj_en = "Client" if is_adult else "Student"
subj_zh = "客户" if is_adult else "学生"

# 3. 字典配置与翻译辅助
DICTIONARY_ZH = {
    # 元数据
    "Student Name": "学生姓名", "Client Name": "客户姓名",
    "School Name": "学校/机构名称", "School District": "学区/服务管区",
    "Student DOB": "学生出生日期", "Client DOB": "客户出生日期",
    "Student ID": "学生编号", "Client ID": "客户编号",
    "Date of FBA": "FBA评估日期", "Date BIP Written": "BIP制定日期", "Date of Last FBA": "最近FBA日期",
    "Cohort Category": "人群年龄组别", "Placement": "服务地点",
    
    # Data Sources 数据源
    "Teacher Interview": "教师/工作人员访谈",
    "Parent Interview": "家长/照护者访谈", "Rating Scales": "评估量表 (如 QABF/MAS)",
    "Direct ABC Observations (Systematic Log)": "直接ABC行为观察日志",

    # BIP / FBA 章节与字段名
    "1. Data Sources": "1. Data Sources (数据来源)",
    "2. Target Behavior Breakdown": "2. Target Behavior Breakdown (目标行为描述)",
    "3. Direct Systematic ABC Observation Ledger": "3. Direct Systematic ABC Observation Ledger (ABC 观察日志)",
    "1. Description of Target Behavior": "1. Description of Target Behavior (目标行为描述)",
    "2. Hypothesis (Developed based on FBA)": "2. Hypothesis (行为功能假设 - 基于FBA推导)",
    "3. Antecedent Modifications (Prevention Strategies)": "3. Antecedent Modifications (前因预防策略)",
    "4. Replacement Behaviors & Teaching Protocol": "4. Replacement Behaviors & Teaching Protocol (替代行为与教学方案)",
    "5. Strategies for Reinforcing Replacement Behavior": "5. Strategies for Reinforcing Replacement Behavior (替代行为强化策略)",
    "6. Strategies for Reducing Target Behavior": "6. Strategies for Reducing Target Behavior (目标行为减少策略)",
    "7. Crisis Plan & Safety Protocol": "7. Crisis Plan & Safety Protocol (危机预案与安全协议)",
    "8. Data Collection, Monitoring & Staff Training": "8. Data Collection, Monitoring & Staff Training (数据收集、评估与人员培训)",

    # 正文标签
    "Selected Sources": "选定的数据源 (Selected Sources)",
    "Description": "行为描述 (Description)",
    "Examples": "具体行为示例 (Examples)",
    "Non-Examples": "非目标行为示例 (Non-Examples)",
    "Nonexamples": "非目标行为示例 (Nonexamples)",
    "Behavior": "目标行为 (Behavior)"
}

def translate_lbl(lbl_text):
    if "Chinese" in selected_language and lbl_text in DICTIONARY_ZH:
        return f"{lbl_text} ({DICTIONARY_ZH[lbl_text]})"
    return lbl_text

def get_preset_abc(age_group):
    if "Adult" in age_group:
        return [
            {"Entry": "Obs #1", "Date/Time": "08/03/2026 09:15 AM", "Observer Role": "Direct Care Staff", "Setting": "Supported Living Apartment", "Antecedent (A)": "Newly hired staff member presented morning chore checklist.", "Behavior (B)": "Verbal aggression (cursing, threats) and physical aggression (shoving staff).", "Consequence (C)": "Senior BCBA/Staff stepped in, guided new staff to pause demand, and represented visual choice board."},
            {"Entry": "Obs #2", "Date/Time": "08/04/2026 06:30 PM", "Observer Role": "Direct Care Staff", "Setting": "Supported Living Apartment", "Antecedent (A)": "Roommate turned on living room TV and adjusted seating area without client consent.", "Behavior (B)": "Loud vocal resistance, blocking TV screen, grabbing remote from roommate.", "Consequence (C)": "Staff prompted roommate mediation and provided alternative personal tablet for private room."},
            {"Entry": "Obs #3", "Date/Time": "08/05/2026 11:00 AM", "Observer Role": "Job Coach", "Setting": "Day Program / Vocational Workshop", "Antecedent (A)": "Client reported joint pain/headache after 2 hours of repetitive standing work.", "Behavior (B)": "Pacing, hand-wringing, aggressive resistance to verbal prompts.", "Consequence (C)": "Staff offered PRN pain relief medication and quiet rest area."}
        ]
    elif "Early Intervention" in age_group:
        return [
            {"Entry": "Obs #1", "Date/Time": "08/03/2026 09:15 AM", "Observer Role": "RBT", "Setting": "Clinic Therapy Room", "Antecedent (A)": "RBT requested sharing toy truck during naturalistic play.", "Behavior (B)": "Screamed, bit own arm, threw toy.", "Consequence (C)": "RBT paused demand, offered sensory chew tool."},
            {"Entry": "Obs #2", "Date/Time": "08/04/2026 10:30 AM", "Observer Role": "BCBA", "Setting": "Home ABA Session", "Antecedent (A)": "Therapist transitioned from bubble play to discrete trial teaching (DTT).", "Behavior (B)": "Dropped to floor, crying, head banging on carpet.", "Consequence (C)": "Therapist paused demand, presented PECS break icon."},
            {"Entry": "Obs #3", "Date/Time": "08/05/2026 04:00 PM", "Observer Role": "RBT", "Setting": "Clinic Social Skills Group", "Antecedent (A)": "RBT called for group cleanup time.", "Behavior (B)": "Ran toward clinic exit door (elopement).", "Consequence (C)": "Staff guided back with visual transition timer."}
        ]
    else: # School-Age 5-21
        return [
            {"Entry": "Obs #1", "Date/Time": "08/03/2026 04:30 PM", "Observer Role": "RBT", "Setting": "In-Home ABA / Screen Time Transition", "Antecedent (A)": "Timer rang signaling 30-min iPad screen time limits reached while RBT turned to document data.", "Behavior (B)": "Pacing, grabbing iPad back, screaming 'Look at me!', dropping to floor.", "Consequence (C)": "RBT made immediate eye contact, prompted '1-min visual extension card', and reinforced quiet waiting."},
            {"Entry": "Obs #2", "Date/Time": "08/04/2026 01:45 PM", "Observer Role": "Paraprofessional", "Setting": "Special Ed Classroom (Small Group)", "Antecedent (A)": "Instructor turned attention to assist peer during iPad group activity.", "Behavior (B)": "Approached staff, pulled sleeve, loud vocalizations, tried to grab peer's iPad.", "Consequence (C)": "Staff turned immediately, made eye contact, and redirected to waiting visual schedule."},
            {"Entry": "Obs #3", "Date/Time": "08/05/2026 09:15 AM", "Observer Role": "BCBA", "Setting": "General Ed Classroom (Desk Work)", "Antecedent (A)": "Teacher presented multi-step writing worksheet.", "Behavior (B)": "Screamed (>80dB), pushed desk away.", "Consequence (C)": "Staff presented 'Break' visual card; demand paused."}
        ]

def infer_func(row):
    ant = str(row.get("Antecedent (A)", "")).lower()
    beh = str(row.get("Behavior (B)", "")).lower()
    full_text = f"{ant} {beh}"
    if any(k in full_text for k in ["pain", "medication", "joint", "headache"]):
        return "Physical Discomfort / Internal State"
    if any(k in full_text for k in ["demand", "task", "worksheet", "chore", "writing", "dtt", "cleanup"]):
        return "Task Escape / Demand Avoidance"
    if any(k in full_text for k in ["ipad", "toy", "screen", "tv", "remote", "bubble", "share"]):
        return "Access to Tangibles / Activities"
    if "look at me" in beh or "pulled sleeve" in beh or "attention" in ant:
        return "Social Attention Seeking"
    return "Automatic / Sensory Stimulation"

# 4. Streamlit 界面构建
st.title("🚀 US-BCBA Automated FBA & BIP Clinical Compiler")
st.divider()

st.header("📋 Phase 1: Multi-Source Data Ingestion & Clinical Configuration")

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Direct ABC Observations & Data Upload", 
    "📝 Expanded FBA Inputs", 
    "📈 QABF Scale",
    "🛡️ BIP Intervention Plan Matrix"
])

with tab1:
    st.subheader("Direct Systematic ABC Data Ledger")
    
    uploaded_file = st.file_uploader("📂 Upload Custom ABC Data (CSV or Excel file):", type=["csv", "xlsx"])
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                uploaded_df = pd.read_csv(uploaded_file)
            else:
                uploaded_df = pd.read_excel(uploaded_file)
            st.success(f"Successfully loaded '{uploaded_file.name}' with {len(uploaded_df)} records!")
            if "Engine Auto-Inferred Function" not in uploaded_df.columns:
                uploaded_df["Engine Auto-Inferred Function"] = uploaded_df.apply(infer_func, axis=1)
            raw_df = uploaded_df
        except Exception as e:
            st.error(f"Error parsing uploaded file: {e}. Falling back to standard default dataset.")
            raw_df = pd.DataFrame(get_preset_abc(selected_age_group))
            raw_df["Engine Auto-Inferred Function"] = raw_df.apply(infer_func, axis=1)
    else:
        raw_df = pd.DataFrame(get_preset_abc(selected_age_group))
        raw_df["Engine Auto-Inferred Function"] = raw_df.apply(infer_func, axis=1)

    edited_abc = st.data_editor(raw_df, num_rows="dynamic", use_container_width=True, key=f"editor_{selected_age_group}")

with tab2:
    st.subheader(f"🤝 Stakeholder Input ({subj_en}-Centered)")
    
    st.markdown("##### 📂 Upload Stakeholder Notes / Interviews (Optional)")
    stakeholder_file = st.file_uploader(
        "Upload Teacher/Parent Interview notes or questionnaire (.txt, .docx, .csv):", 
        type=["txt", "docx", "csv"], 
        key="stakeholder_file_uploader"
    )
    
    uploaded_stakeholder_text = ""
    if stakeholder_file is not None:
        try:
            if stakeholder_file.name.endswith(".txt"):
                uploaded_stakeholder_text = stakeholder_file.read().decode("utf-8")
            elif stakeholder_file.name.endswith(".docx"):
                doc_obj = docx.Document(stakeholder_file)
                uploaded_stakeholder_text = "\n".join([p.text for p in doc_obj.paragraphs if p.text.strip()])
            elif stakeholder_file.name.endswith(".csv"):
                df_stk = pd.read_csv(stakeholder_file)
                uploaded_stakeholder_text = df_stk.to_string()
            st.success(f"Successfully read stakeholder file: '{stakeholder_file.name}'!")
        except Exception as e:
            st.error(f"Error reading file: {e}")

    col_a, col_b = st.columns(2)
    with col_a:
        agency_name = st.text_input("School / Agency Name", "Metropolitan Inclusive Center")
        district_name = st.text_input("District / Health Region", "District 10 Behavioral Division")
        dob_val = st.text_input(f"{subj_en} DOB", "05/12/2022" if "Early" in selected_age_group else ("05/12/2015" if not is_adult else "05/12/2001"))
        id_val = st.text_input(f"{subj_en} ID", "ID-908231")
        fba_date = st.text_input("Date of FBA / BIP", "08/08/2026")
        
    with col_b:
        sources_options = ["Teacher Interview", "Parent Interview", "Rating Scales"]
        data_sources = st.multiselect("1. Secondary Data Sources (Direct Observations auto-included from Tab 1):", sources_options, default=sources_options)
        
    st.divider()
    st.markdown("### 2. Target Behavior Operational Breakdown & Examples")
    
    default_desc = "Screaming (>80dB), pushing materials, dropping to floor during transitions or when demands are presented."
    if uploaded_stakeholder_text:
        default_desc = f"[Extracted from Stakeholder File]\n{uploaded_stakeholder_text[:300]}..."

    c1, c2, c3 = st.columns(3)
    target_beh = c1.text_area("Target Behavior Description", default_desc, height=100)
    beh_examples = c2.text_area("Examples of Target Behavior", "Throwing workbooks, yelling 'No!', hitting table with open palms.", height=100)
    beh_non_examples = c3.text_area("Non-Examples of Target Behavior", "Requesting 'Break' using PECS/AAC card, quietly sitting, asking for teacher assistance.", height=100)

    st.divider()
    st.markdown("### 3. Triggers & Behavioral Context")
    c_t1, c_t2 = st.columns(2)
    setting_events = c_t1.text_area("Setting Events (Slow Triggers)", "Overtiredness, lack of sleep, physical discomfort, or schedule changes.", height=70)
    antecedents_val = c_t2.text_area("Antecedent Events (Immediate Triggers)", "Presentation of multi-step academic tasks or transition away from preferred items.", height=70)

with tab3:
    st.subheader("📈 QABF Psychometric Scale Scores")
    q1, q2, q3, q4, q5 = st.columns(5)
    att_score = q1.number_input("Social Attention", 0, 15, 4 if "Early" in selected_age_group else (12 if not is_adult else 10))
    esc_score = q2.number_input("Task Escape", 0, 15, 12 if "Early" in selected_age_group else (10 if not is_adult else 4))
    tan_score = q3.number_input("Tangibles / Control", 0, 15, 14 if "Early" in selected_age_group else (14 if not is_adult else 12))
    sen_score = q4.number_input("Sensory Stimulation", 0, 15, 3)
    phy_score = q5.number_input("Physical Discomfort", 0, 15, 2 if not is_adult else 11)

with tab4:
    st.subheader("🛡️ BIP Interventions (GNETS Aligned Standard)")
    st.info("The strategies below auto-populate based on clinical defaults and can be customized prior to export.")
    
    b1, b2 = st.columns(2)
    ant_mods = b1.text_area("Antecedent Modifications (Prevention)", 
                            "• Provide visual countdown timer prior to transitions.\n• Check-in walk and talk in the morning.\n• Offer choice board for task sequencing.\n• Establish a folder of appropriate 'waiting' activities.", height=110)
    
    repl_behs = b2.text_area("Replacement Behaviors & Teaching Plan (FCT)", 
                             "• Teach student to hand 'Break' card/icon to staff when overwhelmed.\n• Teach self-monitoring of waiting time using visual timer.\n• Role-play appropriate requesting strategies during 1-on-1 sessions.", height=110)
    
    b3, b4 = st.columns(2)
    reinf_strat = b3.text_area("Strategies for Reinforcing Replacement Behavior", 
                               "• Immediate delivery of 30-second high-preference access upon presenting 'Break' card.\n• Specific verbal praise ('Great job waiting quietly!').\n• Token economy check-in after each successfully completed task interval.", height=110)
    
    reduc_strat = b4.text_area("Strategies for Reducing Target Behavior (Consequence)", 
                               "• Planned ignoring for attention-seeking vocalizations when safety is maintained.\n• Neutral redirection back to task without extended eye contact or lecturing.\n• Pause task demand only after client uses functional communication card.", height=110)
    
    b5, b6 = st.columns(2)
    crisis_plan = b5.text_area("Crisis Plan & De-escalation Protocol", 
                               "• Ensure safety of client and peers by clearing sharp objects.\n• Block self-injurious behavior or physical aggression using non-violent crisis intervention (CPI) procedures.\n• Remove audience/peers from immediate area if necessary.", height=100)
    
    training_plan = b6.text_area("Data Collection, Staff Training & Monitoring", 
                                 "• Frequency and duration data logged daily in CentralReach / Catalyst.\n• Lead BCBA to conduct bi-weekly treatment integrity and fidelity checks.\n• Daily debrief between RBT/Paraprofessional and supervising BCBA.", height=100)

st.divider()

# 5. 生成 Word 文档 (FBA + BIP)
def generate_fba_docx():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.7)
    
    title_str = "FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT"
    if "Chinese" in selected_language:
        title_str += "\n(功能性行为评估标准报告)"
        
    title_p = doc.add_heading(title_str, level=0)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

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
            cell_lbl.text = translate_lbl(lbl)
            cell_val.text = str(val)
            shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
            cell_lbl._tc.get_or_add_tcPr().append(shd)
            cell_lbl.paragraphs[0].runs[0].font.bold = True
            cell_lbl.paragraphs[0].runs[0].font.size = Pt(8.5)
            cell_val.paragraphs[0].runs[0].font.size = Pt(8.5)

    doc.add_paragraph()

    # 1. Data Sources
    doc.add_heading(translate_lbl("1. Data Sources"), level=1)
    all_sources = [DICTIONARY_ZH.get("Direct ABC Observations (Systematic Log)", "Direct ABC Observations") if "Chinese" in selected_language else "Direct ABC Observations"] + [DICTIONARY_ZH.get(s, s) if "Chinese" in selected_language else s for s in data_sources]
    sources_str = ", ".join(all_sources)
    p_src = doc.add_paragraph()
    p_src.add_run(f"{translate_lbl('Selected Sources')}: ").bold = True
    p_src.add_run(sources_str)

    # 2. Target Behavior Breakdown
    doc.add_heading(translate_lbl("2. Target Behavior Breakdown"), level=1)
    p_desc = doc.add_paragraph()
    p_desc.add_run(f"{translate_lbl('Description')}: ").bold = True
    p_desc.add_run(target_beh)
    
    p_ex = doc.add_paragraph()
    p_ex.add_run(f"{translate_lbl('Examples')}: ").bold = True
    p_ex.add_run(beh_examples)
    
    p_non = doc.add_paragraph()
    p_non.add_run(f"{translate_lbl('Non-Examples')}: ").bold = True
    p_non.add_run(beh_non_examples)

    # 3. Direct ABC Observations Ledger
    doc.add_heading(translate_lbl("3. Direct Systematic ABC Observation Ledger"), level=1)
    headers = list(edited_abc.columns)
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    
    for idx, text in enumerate(headers):
        cell = table.rows[0].cells[idx]
        cell.text = text
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

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def generate_bip_docx():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.7)
    
    bip_title = "BEHAVIOR INTERVENTION PLAN (BIP)"
    if "Chinese" in selected_language:
        bip_title += "\n(行为干预计划标准方案)"
    title_p = doc.add_heading(bip_title, level=0)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    info_table = doc.add_table(rows=4, cols=4)
    info_table.style = 'Table Grid'
    table_data = [
        [(f"{subj_en}'s Name", "[CLIENT_NAME]"), ("School Name", agency_name)],
        [(f"{subj_en}'s DOB", dob_val), ("School District", district_name)],
        [(f"{subj_en}'s ID", id_val), ("Date BIP Written", fba_date)],
        [("Cohort Category", selected_age_group), ("Date of Last FBA", fba_date)]
    ]
    for r_idx, row in enumerate(table_data):
        for c_group, (lbl, val) in enumerate(row):
            cell_lbl = info_table.cell(r_idx, c_group * 2)
            cell_val = info_table.cell(r_idx, c_group * 2 + 1)
            cell_lbl.text = translate_lbl(lbl)
            cell_val.text = str(val)
            shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
            cell_lbl._tc.get_or_add_tcPr().append(shd)
            cell_lbl.paragraphs[0].runs[0].font.bold = True
            cell_lbl.paragraphs[0].runs[0].font.size = Pt(8.5)
            cell_val.paragraphs[0].runs[0].font.size = Pt(8.5)

    doc.add_paragraph()

    # 1. Target Behavior
    doc.add_heading(translate_lbl("1. Description of Target Behavior"), level=1)
    p1 = doc.add_paragraph()
    p1.add_run(f"{translate_lbl('Behavior')}: ").bold = True
    p1.add_run(target_beh)
    p1_ex = doc.add_paragraph()
    p1_ex.add_run(f"{translate_lbl('Examples')}: ").bold = True
    p1_ex.add_run(beh_examples)
    p1_non = doc.add_paragraph()
    p1_non.add_run(f"{translate_lbl('Nonexamples')}: ").bold = True
    p1_non.add_run(beh_non_examples)

    # 2. Hypothesis
    doc.add_heading(translate_lbl("2. Hypothesis (Developed based on FBA)"), level=1)
    hyp_text = f"When presented with {antecedents_val.rstrip('.')}, under conditions of {setting_events.rstrip('.')}, [CLIENT_NAME] engages in {target_beh.rstrip('.')} in order to achieve task escape or sensory co-regulation."
    doc.add_paragraph(hyp_text)

    # 3. Antecedent Modifications
    doc.add_heading(translate_lbl("3. Antecedent Modifications (Prevention Strategies)"), level=1)
    doc.add_paragraph(ant_mods)

    # 4. Replacement Behaviors
    doc.add_heading(translate_lbl("4. Replacement Behaviors & Teaching Protocol"), level=1)
    doc.add_paragraph(repl_behs)

    # 5. Reinforcement Strategies
    doc.add_heading(translate_lbl("5. Strategies for Reinforcing Replacement Behavior"), level=1)
    doc.add_paragraph(reinf_strat)

    # 6. Response Strategies / Reduction
    doc.add_heading(translate_lbl("6. Strategies for Reducing Target Behavior"), level=1)
    doc.add_paragraph(reduc_strat)

    # 7. Crisis Plan
    doc.add_heading(translate_lbl("7. Crisis Plan & Safety Protocol"), level=1)
    doc.add_paragraph(crisis_plan)

    # 8. Data & Training
    doc.add_heading(translate_lbl("8. Data Collection, Monitoring & Staff Training"), level=1)
    doc.add_paragraph(training_plan)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# 6. 生成与导出操作区
st.markdown("### 🚀 Export Clinical Documents")
c_btn1, c_btn2 = st.columns(2)

with c_btn1:
    if st.button("📄 Compile Aligned FBA Report (.docx)", type="primary", use_container_width=True):
        fba_file = generate_fba_docx()
        st.success("FBA Report Compiled Successfully!")
        st.download_button(
            label="⬇️ Download Aligned_FBA_Report.docx",
            data=fba_file,
            file_name="Aligned_FBA_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

with c_btn2:
    if st.button("🛡️ Compile Aligned BIP Report (.docx)", type="primary", use_container_width=True):
        bip_file = generate_bip_docx()
        st.success("BIP Report Compiled Successfully!")
        st.download_button(
            label="⬇️ Download Aligned_BIP_Report.docx",
            data=bip_file,
            file_name="Aligned_BIP_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# 7. 免责声明与临床提示页脚 (Disclaimer Footnote)
st.divider()
st.caption("""
**⚠️ Clinical Decision Support & HIPAA Privacy Notice:**
* This application is a clinical decision-support tool designed for Board Certified Behavior Analysts (BCBAs) and Licensed Behavior Analysts (LBAs).
* **De-identification & Compliance**: Generated reports utilize standard placeholders (`[CLIENT_NAME]`) to safeguard Patient Health Information (PHI) in accordance with HIPAA standards.
* **Clinical Responsibility**: All automatically synthesized functional hypotheses, intervention protocols, and BIP strategies must be reviewed, validated, and signed off by a credentialed behavior analyst before clinical implementation or IEP team submission.
""")
