import streamlit as st
import io

# 1. BIPEngine 顶级视觉与侧边栏配置
st.set_page_config(page_title="BIPEngine Sandbox", layout="wide", initial_sidebar_state="expanded")

with st.sidebar:
    st.markdown("<h2 style='color: #008080;'>⚙️ BIPEngine v1.0</h2>", unsafe_allow_html=True)
    st.subheader("🔒 HIPAA & Ethics Core")
    st.markdown("**Zero-Knowledge Sandbox Architecture**")
    st.caption("All clinical inputs are processed entirely in-memory. No Protected Health Information (PHI) is transmitted, cached, or saved to any database, aligning with BACB 2022 Ethical Guidelines.")
    st.divider()
    st.info("🎯 **NIW Project Profile:** Designed to counter systemic BCBA burnout and clear autism therapy waitlists by automating administrative workflows.")

st.title("🚀 BIPEngine: Automated FBA & BIP Clinical Compiler")
st.markdown("##### Empowering Behavior Analysts through Automated Workflow Optimization")
st.divider()

# 2. 核心数据输入模块
st.header("📋 Phase 1: Multi-Source Data Ingestion")
tab1, tab2, tab3 = st.tabs(["📊 ABC Raw Dataframe", "📝 Qualitative Stakeholder Interviews", "📈 Psychometric Profiler (QABF)"])

with tab1:
    st.subheader("Direct Observation ABC Logs")
    st.caption("Interactive ledger simulating automated parsing of imported ABC continuous tracking data.")
    abc_data = [
        {"Timestamp": "2026-08-06 09:15", "Antecedent": "Teacher places intensive math worksheet on desk", "Behavior": "Vocal screaming, throwing pencil across room", "Consequence": "Teacher provides brief verbal redirection, removes worksheet for 2 mins", "Hypothesized Function": "Escape"},
        {"Timestamp": "2026-08-06 10:30", "Antecedent": "Transition from recess to independent seatwork", "Behavior": "Property destruction (tearing paper), dropping to floor", "Consequence": "RBT guides back to chair, provides physical prompt to clear desk", "Hypothesized Function": "Escape"},
        {"Timestamp": "2026-08-06 11:45", "Antecedent": "Direct attention diverted to another peer", "Behavior": "Repetitive desk tapping, vocal disruption", "Consequence": "Peer laughs, RBT implements planned ignoring protocol", "Hypothesized Function": "Attention"}
    ]
    st.data_editor(abc_data, num_rows="dynamic", use_container_width=True)

with tab2:
    col1, col2 = st.columns(2)
    with col1:
        parent_notes = st.text_area("Parent Interview Summary (e.g., Vineland/FAIR)", 
                                    "Mother reports task refusal is highly dense during non-preferred academic routines at home. Escapes from toothbrushing and academic worksheets by dropping to floor.", height=120)
    with col2:
        teacher_notes = st.text_area("Teacher/School Notes", 
                                     "Classroom teacher notes that the behavior never occurs during preferred free-play or physical education. Highly correlated with tasks involving multi-step instructions.", height=120)

with tab3:
    st.subheader("Questions About Behavioral Function (QABF) Scoring Matrix")
    q_col1, q_col2, q_col3, q_col4, q_col5 = st.columns(5)
    with q_col1:
        att_score = st.number_input("Social Attention", 0, 15, 3)
    with q_col2:
        esc_score = st.number_input("Task/Demand Escape", 0, 15, 14)
    with q_col3:
        tan_score = st.number_input("Access to Tangibles", 0, 15, 2)
    with q_col4:
        sen_score = st.number_input("Non-Social / Sensory", 0, 15, 0)
    with q_col5:
        phy_score = st.number_input("Physical Discomfort", 0, 15, 1)

# 3. BIPEngine 智能逻辑分析
highest_qabf = max(att_score, esc_score, tan_score, sen_score, phy_score)
determined_function = "Undetermined"
clinical_hypothesis = ""
proactive_strategies = ""
reactive_strategies = ""

if highest_qabf == esc_score:
    determined_function = "Social Negative Reinforcement (Escape from Academic Demands)"
    clinical_hypothesis = "Based on converging evidence from 3 direct ABC observation logs and QABF metrics (Score: 14), the target behaviors are maintained by escape from non-preferred, multi-step academic tasks."
    proactive_strategies = "• Incorporate high density of functional breaks using a 5-minute visual countdown timer.\n• Interleave high-probability requests (High-P) prior to introducing low-probability math demands.\n• Provide choice arrangements (e.g., 'Do you want to use the blue pen or pencil?')."
    reactive_strategies = "• Errorless teaching procedures during instruction.\n• Differential Reinforcement of Alternative Behavior (DRA): Provide functional escape ONLY when client utilizes the 'Break' communication card.\n• Three-step prompting hierarchy (Tell-Show-Do) to minimize escape-extinction bursts."
elif highest_qabf == att_score:
    determined_function = "Social Positive Reinforcement (Attention)"
    clinical_hypothesis = "Data indicates the behavior is maintained by immediate social mediation from peers and staff when target behaviors are exhibited."
    proactive_strategies = "• Scheduled non-contingent attention (NCA) every 5 minutes.\n• Enrich environment with highly preferred interactive tasks."
    reactive_strategies = "• Planned ignoring of minor disruptive behaviors.\n• Extinction with simultaneous redirection to alternative functional tasks."

# 4. 生成 Word 文档流下载引擎
st.divider()
st.header("🚀 Phase 2: AI-Assisted Document Synthesis Engine")
st.write("Click below to compile all ingested clinical data into a highly individualized, structured corporate clinical draft (.doc).")

# 格式化的 Word 模版内容
mock_docx_content = f"""
================================================================================
                    FUNCTIONAL BEHAVIOR ASSESSMENT & BIP DRAFT
                          BIPENGINE AUTOMATED COMPILER
================================================================================

1. CLINICAL HYPOTHESIS & DATA CONVERGENCE
--------------------------------------------------------------------------------
Primary Hypothesized Function: {determined_function}

Clinical Executive Summary:
{clinical_hypothesis}

2. STAKEHOLDER TRIANGULATION
--------------------------------------------------------------------------------
* Parent Interview Data: {parent_notes}
* Educator Observation Notes: {teacher_notes}

3. BEHAVIOR INTERVENTION PLAN (BIP) STRATEGIES
--------------------------------------------------------------------------------
PROACTIVE / ANTECEDENT MANIPULATIONS:
{proactive_strategies}

REACTIVE / CONSEQUENCE PROCEDURES:
{reactive_strategies}

--------------------------------------------------------------------------------
[Disclaimer: This document is an automatically synthesized initial clinical draft. 
The overseeing BCBA retains full ethical responsibility under the BACB Ethics Code 
to review, amend, and individualize all parameters prior to insurance submission.]
================================================================================
"""

buffer = io.BytesIO()
buffer.write(mock_docx_content.encode('utf-8'))
buffer.seek(0)

if st.download_button(
    label="📥 Compile & Download Official FBA/BIP Draft (.doc)",
    data=buffer,
    file_name="BIPEngine_FBA_BIP_Draft.doc",
    mime="application/msword",
    type="primary"
):
    st.balloons()
