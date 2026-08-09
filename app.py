import streamlit as st
import io
import re
import pandas as pd
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="US-BCBA Clinical BIPEngine v2.7", layout="wide", initial_sidebar_state="expanded")

# 侧边栏
with st.sidebar:
    st.markdown("<h2 style='color: #1F4E78;'>⚙️ BIPEngine v2.7 (US Standard)</h2>", unsafe_allow_html=True)
    st.caption("US BCBA & LBA Clinical Decision Support Engine")
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

is_adult = "Adult" in selected_age_group
subj_en = "Client" if is_adult else "Student"
subj_zh = "客户" if is_adult else "学生"

# 拓展双语字典
DICTIONARY_ZH = {
    "Student Name": "学生姓名", "Client Name": "客户姓名",
    "School/Agency Name": "学校/机构名称", "District / Region": "学区/服务管区",
    "Student DOB": "学生出生日期", "Client DOB": "客户出生日期",
    "Student ID": "学生编号", "Client ID": "客户编号",
    "Date of FBA": "FBA评估日期", "Cohort Category": "人群年龄组别", "Placement Location": "服务地点",
    
    # 干系人访谈列表字段
    "Qualitative Stakeholder Input & Ecological Notes": "质性干系人访谈与环境评估记录",
    "Parent/Caregiver Interview Input": "家长/照护者访谈反馈",
    "Teacher/Staff Interview Input": "教师/工作人员访谈反馈",
    "Student/Client Self-Report Input": "学生/客户自我陈述反馈",
    "Ecological & Environmental Factors": "环境与生态因素评估",
    
    # 常用描述短语
    "Parent reports severe tantrums during homework or screen transitions.": "家长报告情绪发泄常发生于作业时间或电子屏幕切换时。",
    "Teachers/Staff note similar resistance patterns during unstructured group tasks.": "教师/工作人员指出在不具结构的团队活动中也观察到了相似的抗拒模式。",
    "Client expresses frustration with sudden schedule changes and noisy environments.": "客户表达了对突然的日程变更以及噪音环境的挫败感。",
    "Sensory sensitivity to loud background noise and high room temperature noted.": "记录到对背景噪音和高室内温度的感官敏感性。"
}

def translate_text(text_val):
    if not text_val: return ""
    val_str = str(text_val).strip()
    if val_str in DICTIONARY_ZH: return DICTIONARY_ZH[val_str]
    translated = val_str
    for en, zh in DICTIONARY_ZH.items():
        if en in translated: translated = translated.replace(en, zh)
    return translated

# 界面主框架
st.title("🚀 US-BCBA Automated FBA & BIP Clinical Compiler")
st.divider()

st.header("📋 Phase 1: Multi-Source Clinical Ingestion")
tab1, tab2, tab3 = st.tabs([
    "📊 Direct ABC Observations", 
    "📝 Expanded FBA Template Inputs (Stakeholder Input)", 
    "📈 QABF Psychometric Scale"
])

with tab1:
    st.subheader("Direct Systematic ABC Data Ledger")
    # ABC 预设数据 (保持一致)
    raw_df = pd.DataFrame([
        {"Entry": "Obs #1", "Date/Time": "08/03/2026 09:15 AM", "Observer Role": "Direct Care Staff", "Setting": "Clinic Room", "Antecedent (A)": "Presented worksheet.", "Behavior (B)": "Pushed desk.", "Consequence (C)": "Demand paused.", "Engine Auto-Inferred Function": "Task Escape / Demand Avoidance"}
    ])
    edited_abc = st.data_editor(raw_df, num_rows="dynamic", use_container_width=True, key="abc_editor")

with tab2:
    st.subheader(f"🤝 Stakeholder Input & FBA Form Fields ({subj_en}-Centered Format)")
    
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
    st.markdown("### 💬 1.5 Qualitative Stakeholder Input & Ecological Notes")
    st.caption("分角色精细化干系人访谈与生态记录")
    
    # 重新拆分为多条独立的 Stakeholder 输入框
    parent_input = st.text_area(
        "Parent/Caregiver Interview Input (家长/照护者访谈)", 
        "Parent reports severe tantrums during homework or screen transitions.", 
        height=70
    )
    staff_input = st.text_area(
        "Teacher/Staff Interview Input (教师/工作人员访谈)", 
        "Teachers/Staff note similar resistance patterns during unstructured group tasks.", 
        height=70
    )
    student_input = st.text_area(
        f"{subj_en} Self-Report Input ({subj_zh}自我陈述/观察)", 
        f"{subj_en} expresses frustration with sudden schedule changes and noisy environments.", 
        height=70
    )
    ecological_input = st.text_area(
        "Ecological & Environmental Notes (环境生态记录)", 
        "Sensory sensitivity to loud background noise and high room temperature noted.", 
        height=70
    )

    st.divider()
    st.markdown("##### Behavioral Operational Parameters")
    c1, c2, c3 = st.columns(3)
    behavior_desc = c1.text_area("Target Behavior Description", "Verbal threats, shoving staff during transitions.", height=80)
    freq_desc = c2.text_area("Frequency", "3 to 5 episodes per week.", height=80)
    dur_desc = c3.text_area("Duration", "15 to 45 minutes per episode.", height=80)

with tab3:
    st.subheader("📈 QABF Psychometric Scale Scores")
    q1, q2, q3, q4, q5 = st.columns(5)
    att_score = q1.number_input("Social Attention", 0, 15, 3)
    esc_score = q2.number_input("Task Escape", 0, 15, 4)
    tan_score = q3.number_input("Tangibles / Control", 0, 15, 12)
    sen_score = q4.number_input("Sensory Stimulation", 0, 15, 14)
    phy_score = q5.number_input("Physical Discomfort", 0, 15, 11)

# 生成 Word 报告工具函数
def add_bullet_item(doc, title_en, text_en, title_zh):
    p = doc.add_paragraph(style='List Bullet')
    p.add_run(f"{title_en}: ").bold = True
    p.add_run(text_en)
    if "Chinese" in selected_language:
        p.add_run(f"\n（{title_zh}: {translate_text(text_en)}）").italic = True

def generate_fba_document():
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.7)
    
    # 标题
    title_text = f'FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) FORM\n({subj_zh}功能性行为评估标准表)' if "Chinese" in selected_language else 'FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) FORM'
    title = doc.add_heading(title_text, level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 基本信息表格
    header_table = doc.add_table(rows=4, cols=4)
    header_table.style = 'Table Grid'
    info_map = [
        [(f"{subj_en} Name", "[CLIENT_NAME]"), ("School/Agency Name", agency_name)],
        [(f"{subj_en} DOB", dob_val), ("District / Region", district_name)],
        [(f"{subj_en} ID", id_val), ("Date of FBA", fba_date)],
        [("Cohort Category", selected_age_group), ("Placement Location", agency_name)]
    ]
    for r_idx, row_data in enumerate(info_map):
        for c_group, (lbl_en, val) in enumerate(row_data):
            c_lbl = header_table.cell(r_idx, c_group * 2)
            c_val = header_table.cell(r_idx, c_group * 2 + 1)
            c_lbl.text = f"{lbl_en}\n({translate_text(lbl_en)})" if "Chinese" in selected_language else lbl_en
            c_val.text = str(val)
            shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls('w')))
            c_lbl._tc.get_or_add_tcPr().append(shd)
            c_lbl.paragraphs[0].runs[0].font.bold = True
            c_lbl.paragraphs[0].runs[0].font.size = Pt(8.5)
            c_val.paragraphs[0].runs[0].font.size = Pt(8.5)

    doc.add_paragraph()

    # 还原 1.5 Qualitative Stakeholder Input & Ecological Notes 分项列表结构
    sec_title = "1.5 Qualitative Stakeholder Input & Ecological Notes"
    sec_title_zh = "1.5 质性干系人访谈与环境评估记录"
    doc.add_heading(f"{sec_title} ({sec_title_zh})" if "Chinese" in selected_language else sec_title, level=1)
    
    add_bullet_item(doc, "Parent/Caregiver Interview", parent_input, "家长/照护者访谈")
    add_bullet_item(doc, "Teacher/Staff Interview", staff_input, "教师/工作人员访谈")
    add_bullet_item(doc, f"{subj_en} Self-Report", student_input, f"{subj_zh}自我陈述")
    add_bullet_item(doc, "Ecological Notes", ecological_input, "环境生态评估")

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

st.divider()
if st.button("🚀 Compile Perfect FBA (.docx)", type="primary", use_container_width=True):
    fba_file = generate_fba_document()
    st.success("FBA Report Compiled with Restored Bulleted Stakeholder Inputs!")
    st.download_button(
        label="📄 Download Restored_Stakeholder_FBA_Report.docx",
        data=fba_file,
        file_name="Restored_Stakeholder_FBA_Report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
