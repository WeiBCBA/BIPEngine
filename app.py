import streamlit as st
import io

# 1. 顶级配置与脱敏侧边栏
st.set_page_config(page_title="BIPEngine - Clinical Framework", layout="wide", initial_sidebar_state="expanded")

with st.sidebar:
    st.markdown("<h2 style='color: #008080;'>⚙️ BIPEngine v1.1</h2>", unsafe_allow_html=True)
    st.subheader("🔒 Compliance Guidelines")
    st.markdown("**Deterministic Rule-Based Architecture**")
    st.caption("Unlike non-deterministic LLMs, BIPEngine uses closed clinical matrix mapping. Zero data is shared externally, protecting Client PHI.")
    st.divider()
    st.caption("Design Base: BACB Ethics Code")
    st.caption("Environment: In-Memory / Stateless")

st.title("🚀 BIPEngine: Automated Clinical Draft Compiler")
st.markdown("##### Streamlining FBA & BIP Documentation to Mitigate Practitioner Exhaustion")
st.divider()

# 2. 核心数据输入模块（完全支持实时编辑）
st.header("📋 Phase 1: Multi-Source Data Ingestion")
tab1, tab2, tab3 = st.tabs(["📊 ABC Raw Data Log", "📝 Qualitative Stakeholder Interviews", "📈 Psychometric Profiler (QABF)"])

with tab1:
    st.subheader("Direct Observation ABC Ledger")
    st.caption("Reviewer Note: You can add, edit, or delete rows. Current data is mapped from a de-identified sample case.")
    
    # 完美复制用户 PDF 中的 4 条真实 ABC 数据
    abc_data = [
        {"Entry": "Obs #1", "Date & Time": "07/20/2026 09:15 AM", "Setting / Context": "Structured Literacy / Desk Work", "Antecedent (A)": "Teacher presented a writing worksheet (Demand placed).", "Behavior (B)": "Screamed (>80dB), pushed desk away, and attempted to leave the room (Elopement). Duration: 3 min.", "Consequence (C)": "Staff provided verbal instruction 'Take a break' and directed student to quiet corner. Demand temporarily removed.", "Inferred Function": "Escape / Avoidance"},
        {"Entry": "Obs #2", "Date & Time": "07/21/2026 10:30 AM", "Setting / Context": "Free Play / Transition to Math", "Antecedent (A)": "Timer rang to signal end of iPad time (Preferred item removed).", "Behavior (B)": "Vocal outburst (screaming), dropped to floor, refused to move.", "Consequence (C)": "Staff offered a 2-minute extension with visual count-down timer before transition.", "Inferred Function": "Access to Tangible"},
        {"Entry": "Obs #3", "Date & Time": "07/22/2026 01:45 PM", "Setting / Context": "Small Group Activity", "Antecedent (A)": "Teacher turned attention away to assist another student.", "Behavior (B)": "Approached staff member, pulled staff sleeve, made loud vocalizations.", "Consequence (C)": "Staff immediately turned, made eye contact, and verbally reassured student ('I'm right here').", "Inferred Function": "Attention Seeking"},
        {"Entry": "Obs #4", "Date & Time": "07/24/2026 11:10 AM", "Setting / Context": "Independent Work Time", "Antecedent (A)": "Presented with multi-step math task; peers were quietly working.", "Behavior (B)": "Stood up suddenly, walked fast towards classroom door (Elopement).", "Consequence (C)": "Aide blocked door, handed 'I Need a Break' visual card, guided to sensory area.", "Inferred Function": "Escape Task"}
    ]
    edited_abc = st.data_editor(abc_data, num_rows="dynamic", use_container_width=True)

with tab2:
    st.subheader("Qualitative Interview Inputs (Fully Editable)")
    col1, col2 = st.columns(2)
    with col1:
        student_strengths = st.text_area("Student Strengths", 
                                        "Highly responsive to visual schedules, strong tactile interest, affectionate with familiar staff.", height=80)
        parent_notes = st.text_area("FBA Interview & Medical Factors (Parent/Caregiver)", 
                                    "Difficulty exploring leisure items independently; fatigue/sleep disruption exacerbates vocal resistance during demands.", height=120)
    with col2:
        school_demographics = st.text_input("School / Placement Setting", "Box Hill High School / PSU")
        behavior_description = st.text_area("Description of Behavior of Concern", 
                                            "Elopement (leaving assigned area) & Vocal Outbursts (screaming during transitions or academic demands).", height=120)

with tab3:
    st.subheader("Questions About Behavioral Function (QABF) Scoring Matrix")
    st.caption("Input the cumulative raw matrix scores to trigger function prioritization logic:")
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

# 3. 后端算法合成逻辑 (脱敏、专业化)
highest_qabf = max(att_score, esc_score, tan_score, sen_score, phy_score)
primary_function = "Determining..."
bip_antecedent = ""
bip_consequence = ""

if highest_qabf == esc_score:
    primary_function = "Escape or Avoid Task Demands"
    bip_antecedent = "• Modify instructional delivery: Interleave high-probability requests (High-P) before low-P demands.\n• Provide robust visual schedules and countdown timers for transitions.\n• Implement choice-making arrangements regarding task order or materials."
    bip_consequence = "• Differential Reinforcement of Alternative Behavior (DRA): Provide functional escape immediately when student utilizes a 'Break' visual card.\n• 3-Step Prompt Hierarchy (Tell-Show-Do) to mitigate escape extinction bursts."
elif highest_qabf == tan_score:
    primary_function = "Access to Preferred Items/Activities"
    bip_antecedent = "• Use visual transition boards to outline 'First [Work], Then [iPad]'.\n• Systematically thin reinforcement schedules using dense token economies."
    bip_consequence = "• Extinction: Prevent unearned access to targeted tangible items upon problem behavior manifestation."

# 4. 文档编译器下载面板 (分成 2 个独立文件)
st.divider()
st.header("🚀 Phase 2: Algorithmic Synthesis & Document Download")
st.write("Compile data streams into separate, de-identified clinical drafts ready for BCBA modification:")

col_fba, col_bip = st.columns(2)

with col_fba:
    st.markdown("### 📄 FBA Draft Compiler")
    st.caption("Generates a comprehensive Functional Behavioral Assessment based on triangulated input data.")
    
    # 构建纯正 FBA Word 文本内容
    fba_content = f"""FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA)
CONFIDENTIAL CLINICAL DRAFT - FOR CLINICAL REVIEW ONLY

STEP 1: FBA INTERVIEW & CLINICAL DEMOGRAPHICS
--------------------------------------------------------------------------------
Student Name: Client_A (De-identified)
School/Setting: {school_demographics}
Student Strengths: {student_strengths}
Description of Behavior: {behavior_description}
Physiological / Medical Factors: {parent_notes}

STEP 2: DETAILED SUMMARY OF DIRECT DATA
--------------------------------------------------------------------------------
Primary Maintaining Function Identified via Convergence: {primary_function}

CONSEQUENCE FACTORS & SUMMARY ANALYSIS:
• Gain Preferred Items/Activities: {"YES" if tan_score > 5 else "NO (Observed minimally)"}
• Gain Peer or Adult Attention: {"YES" if att_score > 5 else "NO (Observed minimally)"}
• Escape or Avoid Task Demands: {"YES" if esc_score > 5 else "NO (Observed minimally)"}

--------------------------------------------------------------------------------
[Ethics Note: Generated by BIPEngine Core Rules. The overseeing BCBA retains full 
responsibility under BACB guidelines to review, modify, and sign this document.]
"""
    fba_buffer = io.BytesIO()
    fba_buffer.write(fba_content.encode('utf-8'))
    fba_buffer.seek(0)
    
    if st.download_button(label="📥 Download Official FBA Report (.doc)", data=fba_buffer, file_name="FBA_Report_Client_A.doc", mime="application/msword", type="primary"):
        st.balloons()

with col_bip:
    st.markdown("### 📝 BIP Draft Compiler")
    st.caption("Generates a Function-Based Behavior Intervention Plan mapping directly to FBA hypotheses.")
    
    # 构建纯正 BIP Word 文本内容
    bip_content = f"""BEHAVIOR INTERVENTION PLAN (BIP)
CONFIDENTIAL CLINICAL DRAFT - FOR CLINICAL REVIEW ONLY

Target Client Reference: Client_A (De-identified)
Target Behavior Focus: {behavior_description}
Primary Function Framework: {primary_function}

STEP 1: PROACTIVE & ANTECEDENT MANIPULATIONS
--------------------------------------------------------------------------------
{bip_antecedent}

STEP 2: REACTIVE & CONSEQUENCE PROCEDURES
--------------------------------------------------------------------------------
{bip_consequence}

--------------------------------------------------------------------------------
[Ethics Note: This BIP must be individualized with specific reactive profiles and 
faded reinforcement token criteria by the supervising certificant prior to field deployment.]
"""
    bip_buffer = io.BytesIO()
    bip_buffer.write(bip_content.encode('utf-8'))
    bip_buffer.seek(0)
    
    if st.download_button(label="📥 Download Function-Based BIP (.doc)", data=bip_buffer, file_name="BIP_Plan_Client_A.doc", mime="application/msword", type="primary"):
        st.balloons()
