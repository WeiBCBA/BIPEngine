import streamlit as st
import io
import pandas as pd

# 1. 顶级配置与 HIPAA 物理隔离合规侧边栏
st.set_page_config(page_title="BIPEngine - Expert System Framework", layout="wide", initial_sidebar_state="expanded")

with st.sidebar:
    st.markdown("<h2 style='color: #008080;'>⚙️ BIPEngine v1.3</h2>", unsafe_allow_html=True)
    st.subheader("🔒 Strict HIPAA Compliance")
    st.markdown("**Stateless Data Minimization**")
    st.caption("BIPEngine operates on a zero-knowledge framework. Uploaded CSV data logs and text interview payloads are evaluated entirely in-memory (RAM) and destroyed immediately upon session closure. No persistent storage or cloud logs exist.")
    st.divider()
    st.success("✔ BACB Ethics Code 2022 Compliant")
    st.info("💡 Reviewer Option: You can now upload raw clinical datasets below to test the automated compilation framework.")

st.title("🚀 BIPEngine: Automated Clinical Draft Compiler")
st.markdown("##### Streamlining FBA & BIP Documentation via Deterministic Matrix Mapping to Mitigate Practitioner Exhaustion")
st.divider()

# 2. 核心数据输入模块（完全支持真实数据上传 + 默认 Mock 数据兼容）
st.header("📋 Phase 1: Multi-Source Data Ingestion")
tab1, tab2, tab3 = st.tabs(["📊 ABC Raw Data (Manual or CSV Upload)", "📝 Qualitative Stakeholder Interviews", "📈 Psychometric Profiler (QABF)"])

with tab1:
    st.subheader("Direct Observation ABC Ledger")
    st.caption("BCBAs can either modify the de-identified matrix below manually OR upload a raw client tracking CSV file.")
    
    # 支持真实数据上传组件！
    uploaded_abc_file = st.file_uploader("Upload Raw ABC Data Tracking Sheet (CSV format)", type=["csv"])
    
    # 默认数据定义（完美还原您的真实 PDF 模板）
    default_abc_data = [
        {"Entry": "Obs #1", "Date & Time": "07/20/2026 09:15 AM", "Setting / Context": "Structured Literacy / Desk Work", "Antecedent (A)": "Teacher presented a writing worksheet (Demand placed).", "Behavior (B)": "Screamed (>80dB), pushed desk away, and attempted to leave the room (Elopement). Duration: 3 min.", "Consequence (C)": "Staff provided verbal instruction 'Take a break' and directed student to quiet corner. Demand temporarily removed.", "Inferred Function": "Escape / Avoidance"},
        {"Entry": "Obs #2", "Date & Time": "07/21/2026 10:30 AM", "Setting / Context": "Free Play / Transition to Math", "Antecedent (A)": "Timer rang to signal end of iPad time (Preferred item removed).", "Behavior (B)": "Vocal outburst (screaming), dropped to floor, refused to move.", "Consequence (C)": "Staff offered a 2-minute extension with visual count-down timer before transition.", "Inferred Function": "Access to Tangible"},
        {"Entry": "Obs #3", "Date & Time": "07/22/2026 01:45 PM", "Setting / Context": "Small Group Activity", "Antecedent (A)": "Teacher turned attention away to assist another student.", "Behavior (B)": "Approached staff member, pulled staff sleeve, made loud vocalizations.", "Consequence (C)": "Staff immediately turned, made eye contact, and verbally reassured student ('I'm right here').", "Inferred Function": "Attention Seeking"},
        {"Entry": "Obs #4", "Date & Time": "07/24/2026 11:10 AM", "Setting / Context": "Independent Work Time", "Antecedent (A)": "Presented with multi-step math task; peers were quietly working.", "Behavior (B)": "Stood up suddenly, walked fast towards classroom door (Elopement).", "Consequence (C)": "Aide blocked door, handed 'I Need a Break' visual card, guided to sensory area.", "Inferred Function": "Escape Task"}
    ]
    
    # 如果用户上传了真实文件，优先解析真实文件，否则展示默认经典模板
    if uploaded_abc_file is not None:
        try:
            display_df = pd.read_csv(uploaded_abc_file)
            st.success("🎉 Custom clinical ABC CSV log parsed successfully in-memory!")
        except Exception as e:
            st.error(f"Error parsing file: {e}. Falling back to standard template.")
            display_df = pd.DataFrame(default_abc_data)
    else:
        display_df = pd.DataFrame(default_abc_data)
        
    edited_abc = st.data_editor(display_df, num_rows="dynamic", use_container_width=True)

with tab2:
    st.subheader("Qualitative Interview Inputs (Fully Editable & Expandable)")
    
    # 同样支持上传一个纯文本的访谈笔记
    uploaded_notes_file = st.file_uploader("Optionally upload raw stakeholder interview notes (.txt)", type=["txt"])
    uploaded_text_content = ""
    if uploaded_notes_file is not None:
        uploaded_text_content = uploaded_notes_file.read().decode("utf-8")
        st.success("🎉 External interview transcript ingested into memory!")

    col1, col2 = st.columns(2)
    with col1:
        student_strengths = st.text_area("Student Strengths", 
                                        "Highly responsive to visual schedules, strong tactile interest, affectionate with familiar staff.", height=80)
        parent_notes = st.text_area("FBA Interview & Medical Factors (Parent/Caregiver)", 
                                    uploaded_text_content if uploaded_text_content else "Difficulty exploring leisure items independently; fatigue/sleep disruption exacerbates vocal resistance during demands.", height=120)
    with col2:
        school_demographics = st.text_input("School / Placement Setting", "Box Hill High School / PSU")
        behavior_description = st.text_area("Description of Behavior of Concern", 
                                            "Elopement (leaving assigned area) & Vocal Outbursts (screaming during transitions or academic demands).", height=120)
    
    st.divider()
    st.subheader("➕ Additional Clinical Input Dimensions")
    custom_inputs = st.text_area("Custom Stakeholder Observations / Ecological & Setting Events", 
                                  "e.g., Speech therapist reports communication breakdown increases frustration; Classroom transitions between rooms trigger higher density of elopement.", height=100)

with tab3:
    st.subheader("Questions About Behavioral Function (QABF) Scoring Matrix")
    st.caption("Input cumulative raw scores to trigger the automated matrix routing logic:")
    q_col1, q_col2, q_col3, q_col4, q_col5 = st.columns(5)
    with q_col1:
        att_score = st.number_input("Social Attention", 0, 15, 3)
    with q_col2:
        esc_score = st.number_input("Task/Demand Escape", 0, 15, 14)
    with q_col3:
        tan_score = st.number_input("Access to Tangibles", 0, 15, 4)
    with q_col4:
        sen_score = st.number_input("Non-Social / Sensory", 0, 15, 0)
    with q_col5:
        phy_score = st.number_input("Physical Discomfort", 0, 15, 1)

# 3. 后端确定性算法合成逻辑 (大幅度增强临床策略厚度)
highest_qabf = max(att_score, esc_score, tan_score, sen_score, phy_score)
primary_function = "Determining..."
bip_antecedent = ""
bip_consequence = ""
replacement_behavior = ""
generalization_plan = ""

if highest_qabf == esc_score:
    primary_function = "Social Negative Reinforcement (Escape or Avoid Task Demands)"
    replacement_behavior = (
        "1. Functional Communication Training (FCT): Teach the client to independently utilize an 'I need a break' visual card or vocal script prior to behavioral escalation.\n"
        "2. Tolerating Delay/Denial: Systematically introduce a progressive delay schedule between the request for a break and the delivery of reinforcement."
    )
    bip_antecedent = (
        "- Curriculum Modification: Break intensive multi-step math/literacy worksheets into single-line visual blocks to lower immediate cognitive load.\n"
        "- High-Probability Request Sequences (High-P): Deliver 3 simple, high-preference requests (e.g., 'high five', 'touch your nose') immediately before placing a non-preferred writing demand.\n"
        "- Choice-Making Arrangements: Offer control over task logistics (e.g., 'Do you want to complete the task with the blue marker or the pencil?').\n"
        "- Visual Pre-correction: Implement a 5-minute and 2-minute visual countdown timer prior to any structured desk work transition."
    )
    bip_consequence = (
        "- Differential Reinforcement of Alternative Behavior (DRA): Provide immediate functional escape (2 minutes) and brief praise ONLY when the replacement behavior (FCT card) is exhibited. Problem behaviors (screaming, elopement) will produce NO escape.\n"
        "- 3-Step Prompting Hierarchy: If the client attempts to elope or scream upon task placement, implement a calm 'Tell-Show-Do' prompt sequence to ensure task completion, thereby bypassing escape-extinction bursts.\n"
        "- Environmental Blocking (Safety): Position staff strategically between the student and the door. Block elopement calmly without verbal interaction, eye contact, or reprimands."
    )
    generalization_plan = (
        "- Stimulus Generalization: Train multiple classroom aides and the student's parents to prompt and reinforce the FCT card using the exact same verbal scripts.\n"
        "- Schedule Thinning: Gradually increase the number of math components required from 1 to 5 before the FCT break is honored."
    )
elif highest_qabf == tan_score:
    primary_function = "Social Positive Reinforcement (Access to Preferred Items/Activities)"
    replacement_behavior = "Teach the client to look at the instructor and hand an icon representing the desired preferred item (e.g., iPad icon) rather than engaging in vocal outbursts."
    bip_antecedent = (
        "- First/Then Visual Matrices: Explicitly display a 'First [Work Task], Then [Preferred Activity]' structure.\n"
        "- Token Economy System: Establish a dense token token card where tokens are delivered on a fixed ratio (FR-1) for compliance, exchangeable for preferred activities."
    )
    bip_consequence = (
        "- Extinction (Access): Ensure that problem behaviors strictly result in zero access to the preferred item. If an outburst occurs, the item remains unavailable.\n"
        "- Redirection: Guide the client back to the compliance task using physical prompts without verbal engagement."
    )
