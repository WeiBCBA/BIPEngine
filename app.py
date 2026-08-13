import io
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
import pandas as pd
import plotly.express as px
import streamlit as st

# ==========================================
# 1. Page Configuration & Custom CSS
# ==========================================
st.set_page_config(
    page_title="Full Lifecycle Special Education SandBox (全生命周期特教沙盒)",
    page_icon="🧩",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling for Clinical Dashboard
st.markdown(
    """
    <style>
    .main-header { font-size: 2.2rem; color: #1F4E78; font-weight: 700; margin-bottom: 0.2rem; }
    .sub-header { font-size: 1.0rem; color: #555; margin-bottom: 1.5rem; }
    .privacy-banner {
        background-color: #EBF3FA;
        border-left: 5px solid #1F4E78;
        padding: 1rem;
        border-radius: 4px;
        margin-bottom: 1.5rem;
    }
    .cohort-card {
        background-color: #F8F9FA;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 1.2rem;
        margin-bottom: 1rem;
    }
    .stButton>button { border-radius: 6px; font-weight: 600; }
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. Encapsulated Mock Golden Datasets
# ==========================================
MOCK_DATASETS = {
    "group1": {
        "cohort_name": "Early Intervention Protocol (2-5 Years Old)",
        "framework": "ESDM / NDBI (Early Start Denver Model)",
        "client_meta": {
            "name": "[CLIENT_NAME]",
            "age": "3.5 yrs",
            "dob": "02/14/2023",
            "id": "EI-3092",
            "setting": "Early Childhood Clinic & Home ABA",
            "agency": "Early Developmental Intervention Center",
            "district": "Region 1 Early Intervention Network",
        },
        "qabf_scores": {
            "Social Attention": 3,
            "Task Escape": 5,
            "Tangible Access": 6,
            "Sensory / Automatic": 14,
            "Physical Discomfort": 9,
        },
        "target_behavior": (
            "Severe temper tantrums (screaming, body stiffness), PICA"
            " (ingesting sand, non-edible foam, raw dough), and self-injurious"
            " head-banging on carpet."
        ),
        "examples": (
            "Slapping face when loud noises occur; picking up sand during"
            " playground transition and attempting to swallow; dropping to floor"
            " crying during DTT transitions."
        ),
        "non_examples": (
            "Reaching for chewy tool, pointing to PECS icon for 'Break',"
            " accepting co-regulation hug."
        ),
        "antecedents": (
            "Sudden sensory overstimulation (loud kitchen noises, unexpected"
            " touch), transitions from unstructured sensory play to table-top"
            " DTT tasks."
        ),
        "setting_events": (
            "Teething pain, gastrointestinal discomfort, disrupted nap time."
        ),
        "bip_strategies": {
            "antecedent": (
                "• Implement NDBI/ESDM joint action routines.\n• Provide"
                " ambient noise-canceling headphones in loud settings.\n• Use"
                " visual transition countdown cards with sensory chew tools."
            ),
            "replacement": (
                "• Teach Functional Communication (FCT): handing PECS 'More' or"
                " 'Break' icon.\n• Teach joint attention signaling (pointing to"
                " desired sensory items)."
            ),
            "reinforcement": (
                "• Immediate 30-second access to high-preference tactile"
                " sensory toys upon communication card exchange.\n• Enthusiastic"
                " social praise coregulation."
            ),
            "reduction": (
                "• Block PICA attempts neutrally using physical redirection.\n•"
                " Cushion head-banging area with soft foam pad without verbal"
                " lecturing or extended eye contact."
            ),
            "crisis": (
                "• Maintain clear radius around client.\n• Use physical blocks"
                " for SIB.\n• Contact parents if physical distress/fever is"
                " suspected."
            ),
            "training": (
                "• RBTs and parents trained on ESDM fidelity checklists weekly"
                " by BCBA."
            ),
        },
        "abc_data": pd.DataFrame([
            {
                "Entry": "Obs #1",
                "Date/Time": "08/10/2026 09:15 AM",
                "Setting": "Clinic Playroom",
                "Antecedent (A)": (
                    "Kitchen blender noise started in adjacent breakroom"
                ),
                "Behavior (B)": "Screamed, slapped face 3x, dropped to floor",
                "Consequence (C)": (
                    "RBT offered noise-canceling headphones and sensory chew tool"
                ),
                "Engine Auto-Inferred Function": "Automatic / Sensory",
            },
            {
                "Entry": "Obs #2",
                "Date/Time": "08/10/2026 10:30 AM",
                "Setting": "Outdoor Sandbox",
                "Antecedent (A)": (
                    "Transition prompt given to pack up sand toys"
                ),
                "Behavior (B)": (
                    "Grabbed handful of sand and attempted to put in mouth"
                ),
                "Consequence (C)": (
                    "RBT blocked mouth, redirected to oral chew tool, demand"
                    " paused"
                ),
                "Engine Auto-Inferred Function": "Automatic / Sensory",
            },
            {
                "Entry": "Obs #3",
                "Date/Time": "08/11/2026 11:00 AM",
                "Setting": "Table Therapy Room",
                "Antecedent (A)": "Presented matching discrete trial worksheet",
                "Behavior (B)": "Head banging on foam floor mat",
                "Consequence (C)": (
                    "Therapist paused demand, presented PECS 'Break' card"
                ),
                "Engine Auto-Inferred Function": "Task Escape / Demand Avoidance",
            },
        ]),
    },
    "group2": {
        "cohort_name": "School-Age / IEP Protocol (5-21 Years Old)",
        "framework": "IDEA / IEP / PBIS (Individuals with Disabilities Act)",
        "client_meta": {
            "name": "[CLIENT_NAME]",
            "age": "11.2 yrs",
            "dob": "05/12/2015",
            "id": "IEP-8821",
            "setting": "Special Education Classroom & Resource Room",
            "agency": "Metropolitan Inclusive School District",
            "district": "Suburban Special Ed Co-Op District 10",
        },
        "qabf_scores": {
            "Social Attention": 11,
            "Task Escape": 15,
            "Tangible Access": 7,
            "Sensory / Automatic": 2,
            "Physical Discomfort": 3,
        },
        "target_behavior": (
            "Classroom disruption, vocal aggression (>80dB screaming),"
            " throwing academic workbooks, and desk-pushing during independent"
            " work."
        ),
        "examples": (
            "Yelling 'No way, I'm not doing this!', sweeping paper/pencils off"
            " desk, pushing chair backwards into aisle."
        ),
        "non_examples": (
            "Raising hand for teacher assistance, placing 'Break Card' on desk,"
            " quietly sitting."
        ),
        "antecedents": (
            "Presentation of multi-step independent writing tasks, peer"
            " distractions, reduction of direct teacher attention."
        ),
        "setting_events": (
            "Poor sleep night prior (<6 hours logged), morning bus dispute."
        ),
        "bip_strategies": {
            "antecedent": (
                "• Chunk writing assignments into 3-step manageable cards.\n•"
                " Pre-teach task vocabulary before group work.\n• Provide visual"
                " schedule checklist on desk."
            ),
            "replacement": (
                "• Teach student to hand 'Help Card' or '1-Min Break' icon to"
                " instructor.\n• Teach self-graphing of task completion."
            ),
            "reinforcement": (
                "• Immediate deliver 2-minute preferred computer draw time upon"
                " using 'Help/Break Card'.\n• Token economy check-in every 15"
                " minutes."
            ),
            "reduction": (
                "• Planned ignoring for low-level vocal grumbles.\n• Neutral"
                " redirection back to visual checklist without extended verbal"
                " reprimands."
            ),
            "crisis": (
                "• Ensure peer safety; guide classroom peers to secondary"
                " room if physical desk-tipping occurs.\n• Follow district CPI"
                " de-escalation guidelines."
            ),
            "training": (
                "• Classroom staff and paraprofessionals complete bi-weekly"
                " treatment fidelity scoring."
            ),
        },
        "abc_data": pd.DataFrame([
            {
                "Entry": "Obs #1",
                "Date/Time": "08/10/2026 09:30 AM",
                "Setting": "Gen-Ed Classroom",
                "Antecedent (A)": (
                    "Teacher presented 2-page independent math worksheet"
                ),
                "Behavior (B)": "Screamed 'I won't do it!', pushed desk away",
                "Consequence (C)": (
                    "Staff presented 'Break' visual card; demand paused for 2"
                    " minutes"
                ),
                "Engine Auto-Inferred Function": "Task Escape / Demand Avoidance",
            },
            {
                "Entry": "Obs #2",
                "Date/Time": "08/11/2026 01:15 PM",
                "Setting": "Small Group Reading",
                "Antecedent (A)": "Teacher turned attention to help peer",
                "Behavior (B)": (
                    "Threw textbook across desk, shouted 'Look at me!'"
                ),
                "Consequence (C)": (
                    "Staff redirected with neutral tone to waiting visual"
                    " schedule"
                ),
                "Engine Auto-Inferred Function": "Social Attention Seeking",
            },
            {
                "Entry": "Obs #3",
                "Date/Time": "08/12/2026 10:45 AM",
                "Setting": "Resource Room",
                "Antecedent (A)": "Multi-step essay prompt assigned",
                "Behavior (B)": "Tore paper, swept markers onto floor",
                "Consequence (C)": (
                    "Guided to quiet break area; task chunked into single"
                    " sentence prompt"
                ),
                "Engine Auto-Inferred Function": "Task Escape / Demand Avoidance",
            },
        ]),
    },
    "group3": {
        "cohort_name": "Adult Community / NDIS Lifespan Protocol (21+ Years)",
        "framework": (
            "NDIS PBS (Positive Behaviour Support) & Restrictive Practice"
            " Reduction"
        ),
        "client_meta": {
            "name": "[CLIENT_NAME]",
            "age": "26.8 yrs",
            "dob": "11/04/1999",
            "id": "NDIS-90123",
            "setting": "Supported Independent Living (SIL) Apartment & Day Vo-Tech",
            "agency": "Lifespan Adult Community Care",
            "district": "Metro Disability Services Region",
        },
        "qabf_scores": {
            "Social Attention": 5,
            "Task Escape": 4,
            "Tangible Access": 14,
            "Sensory / Automatic": 12,
            "Physical Discomfort": 4,
        },
        "target_behavior": (
            "Property destruction (knocking over vocational assembly tables,"
            " tearing community center curtains), and verbal aggression towards"
            " support staff."
        ),
        "examples": (
            "Sweeping items off breakroom counters, shouting threats, blocking"
            " hallway access during room transitions."
        ),
        "non_examples": (
            "Verbally stating 'I need quiet time', using personal tablet to show"
            " choice card to key worker."
        ),
        "antecedents": (
            "Unannounced changes in daily vocational schedule, delayed access to"
            " preferred personal tablet/TV, shared living room conflicts."
        ),
        "setting_events": (
            "Medication transition phase, noisy roommate environment."
        ),
        "bip_strategies": {
            "antecedent": (
                "• Co-design daily schedule every morning using personal iPad"
                " planner.\n• Ensure 15-minute advance notification before any"
                " staff rotation.\n• Provide private single-occupancy quiet"
                " room."
            ),
            "replacement": (
                "• Self-advocacy training: Teach client to state 'I want my"
                " space/tablet now'.\n• Independent self-management checklist."
            ),
            "reinforcement": (
                "• Immediate access to preferred community outings/activities"
                " upon self-advocacy communication.\n• Monthly person-centered"
                " goal incentive rewards."
            ),
            "reduction": (
                "• Maintain safe distance during property destruction without"
                " physical restraint.\n• Zero restrictive practices without"
                " emergency authorization."
            ),
            "crisis": (
                "• Clear area of bystanders.\n• Follow NDIS Quality and"
                " Safeguards Commission emergency reporting protocol."
            ),
            "training": (
                "• All SIL support workers complete person-centered active"
                " support (PCAS) training."
            ),
        },
        "abc_data": pd.DataFrame([
            {
                "Entry": "Obs #1",
                "Date/Time": "08/09/2026 09:00 AM",
                "Setting": "SIL Apartment",
                "Antecedent (A)": (
                    "New support worker introduced schedule change"
                ),
                "Behavior (B)": "Swept dishes off table, shouted threats",
                "Consequence (C)": (
                    "Senior staff stepped in, offered visual choice board,"
                    " demand paused"
                ),
                "Engine Auto-Inferred Function": "Access to Tangibles / Control",
            },
            {
                "Entry": "Obs #2",
                "Date/Time": "08/10/2026 02:00 PM",
                "Setting": "Vocational Workshop",
                "Antecedent (A)": "Tablet time limit reached during break",
                "Behavior (B)": (
                    "Pacing, grabbed tablet back, knocked over assembly chair"
                ),
                "Consequence (C)": (
                    "Job coach prompted self-advocacy phrase card and granted"
                    " 5-min extension"
                ),
                "Engine Auto-Inferred Function": "Access to Tangibles / Control",
            },
            {
                "Entry": "Obs #3",
                "Date/Time": "08/11/2026 06:30 PM",
                "Setting": "Community Living Room",
                "Antecedent (A)": (
                    "Roommate adjusted TV channel without consensus"
                ),
                "Behavior (B)": "Blocked TV screen, loud vocal resistance",
                "Consequence (C)": (
                    "Staff facilitated structured roommate mediation"
                ),
                "Engine Auto-Inferred Function": "Access to Tangibles / Control",
            },
        ]),
    },
}

# ==========================================
# 3. Session State Initialization
# ==========================================
if "active_group" not in st.session_state:
  st.session_state.active_group = "group1"
if "loaded_data" not in st.session_state:
  st.session_state.loaded_data = MOCK_DATASETS["group1"]
if "is_custom_uploaded" not in st.session_state:
  st.session_state.is_custom_uploaded = False

# ==========================================
# 4. Top Privacy & HIPAA Notice Banner
# ==========================================
st.markdown(
    """
    <div class="privacy-banner">
        <h3 style="margin-top:0; color:#1F4E78;">🔒 Zero-Cloud Security & HIPAA Compliance Notice</h3>
        <p style="margin-bottom:0.3rem;">
            <strong>100% On-Premise Memory Processing:</strong> No client data or uploaded documents are stored in cloud databases. All session memory is immediately wiped upon tab close.
        </p>
        <p style="margin-bottom:0;">
            💡 <em>To guarantee 100% HIPAA compliance, you can either upload your own fully de-identified (.txt, .md, .docx) case notes OR simply click the <strong>'Auto-Load Sample Dataset'</strong> button in any cohort tab to test our standard golden data suite.</em>
        </p>
    </div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    "<div class='main-header'>🧩 Full Lifecycle Special Education SandBox (全生命周期特教沙盒)</div>",
    unsafe_allow_html=True,
)
st.markdown(
    "<div class='sub-header'>Multi-Cohort Clinical Compliance Decision Engine"
    " (ESDM • IEP • NDIS Lifespan)</div>",
    unsafe_allow_html=True,
)


# Helper function to switch datasets
def load_cohort_dataset(group_key, custom_text=None):
  st.session_state.active_group = group_key
  base_dataset = MOCK_DATASETS[group_key].copy()

  if custom_text:
    base_dataset["target_behavior"] = (
        f"[Extracted from Custom Notes]\n{custom_text[:300]}..."
    )
    st.session_state.is_custom_uploaded = True
  else:
    st.session_state.is_custom_uploaded = False

  st.session_state.loaded_data = base_dataset


# ==========================================
# 5. Cohort Sandbox Tabs
# ==========================================
tab_g1, tab_g2, tab_g3 = st.tabs([
    "👶 Group 1: Early Intervention (2-5 Yrs)",
    "🏫 Group 2: School-Age / IEP (5-21 Yrs)",
    "💼 Group 3: Adult Lifespan / NDIS (21+ Yrs)",
])

# ----- Group 1 Tab -----
with tab_g1:
  st.markdown("### 👶 Group 1: 2-5 Years Old (Early Intervention Protocol)")
  st.caption(
      "Clinical Standard: ESDM (Early Start Denver Model) / NDBI"
      " Developmental Compliance"
  )

  col_u1, col_b1 = st.columns([2, 1])
  with col_u1:
    up1 = st.file_uploader(
        "Upload Case Notes / ABC Logs (.txt, .md, .docx, .csv):",
        type=["txt", "md", "docx", "csv"],
        key="uploader_g1",
    )
    if up1:
      try:
        content = (
            up1.read().decode("utf-8")
            if not up1.name.endswith(".docx")
            else "Docx loaded"
        )
        load_cohort_dataset("group1", custom_text=content)
        st.success(f"Loaded custom file for Group 1: '{up1.name}'")
      except Exception as e:
        st.error(f"Error parsing file: {e}")

  with col_b1:
    st.write(" ")
    st.write(" ")
    if st.button(
        "💡 Auto-Load Pre-cleansed Sample Dataset #1",
        use_container_width=True,
        type="primary" if st.session_state.active_group == "group1" else "secondary",
    ):
      load_cohort_dataset("group1")
      st.success("Group 1 Golden Dataset Loaded!")

# ----- Group 2 Tab -----
with tab_g2:
  st.markdown("### 🏫 Group 2: 5-21 Years Old (School-Age / IEP Protocol)")
  st.caption(
      "Clinical Standard: IDEA IEP Accommodations, PBIS & Functional"
      " Communication Training (FCT)"
  )

  col_u2, col_b2 = st.columns([2, 1])
  with col_u2:
    up2 = st.file_uploader(
        "Upload Case Notes / ABC Logs (.txt, .md, .docx, .csv):",
        type=["txt", "md", "docx", "csv"],
        key="uploader_g2",
    )
    if up2:
      try:
        content = (
            up2.read().decode("utf-8")
            if not up2.name.endswith(".docx")
            else "Docx loaded"
        )
        load_cohort_dataset("group2", custom_text=content)
        st.success(f"Loaded custom file for Group 2: '{up2.name}'")
      except Exception as e:
        st.error(f"Error parsing file: {e}")

  with col_b2:
    st.write(" ")
    st.write(" ")
    if st.button(
        "💡 Auto-Load Pre-cleansed Sample Dataset #2",
        use_container_width=True,
        type="primary" if st.session_state.active_group == "group2" else "secondary",
    ):
      load_cohort_dataset("group2")
      st.success("Group 2 Golden Dataset Loaded!")

# ----- Group 3 Tab -----
with tab_g3:
  st.markdown("### 💼 Group 3: 21+ Years Old (Adult Community / NDIS Lifespan)")
  st.caption(
      "Clinical Standard: NDIS PBS (Positive Behaviour Support), Restrictive"
      " Practice Reduction & Independent Living"
  )

  col_u3, col_b3 = st.columns([2, 1])
  with col_u3:
    up3 = st.file_uploader(
        "Upload Case Notes / ABC Logs (.txt, .md, .docx, .csv):",
        type=["txt", "md", "docx", "csv"],
        key="uploader_g3",
    )
    if up3:
      try:
        content = (
            up3.read().decode("utf-8")
            if not up3.name.endswith(".docx")
            else "Docx loaded"
        )
        load_cohort_dataset("group3", custom_text=content)
        st.success(f"Loaded custom file for Group 3: '{up3.name}'")
      except Exception as e:
        st.error(f"Error parsing file: {e}")

  with col_b3:
    st.write(" ")
    st.write(" ")
    if st.button(
        "💡 Auto-Load Pre-cleansed Sample Dataset #3",
        use_container_width=True,
        type="primary" if st.session_state.active_group == "group3" else "secondary",
    ):
      load_cohort_dataset("group3")
      st.success("Group 3 Golden Dataset Loaded!")

st.divider()

# ==========================================
# 6. Dynamic Processing Engine & Dashboard
# ==========================================
current_data = st.session_state.loaded_data

st.markdown(f"## ⚡ Active Case Engine: {current_data['cohort_name']}")
st.info(f"**Clinical Compliance Framework:** {current_data['framework']}")

col_left, col_right = st.columns([1, 1])

# --- Left Column: Interactive QABF Chart ---
with col_left:
  st.markdown("### 📈 Dynamic QABF Psychometric Profile")
  scores = current_data["qabf_scores"]
  df_qabf = pd.DataFrame(
      {"Behavioral Function": list(scores.keys()), "Score": list(scores.values())}
  )

  fig = px.bar(
      df_qabf,
      x="Behavioral Function",
      y="Score",
      color="Score",
      color_continuous_scale="Blues",
      text="Score",
      title=f"QABF Subscale Analysis ({current_data['client_meta']['age']})",
  )
  fig.update_layout(yaxis_range=[0, 15], showlegend=False, height=350)
  st.plotly_chart(fig, use_container_width=True)

# --- Right Column: ABC Summary & Trend Analysis ---
with col_right:
  st.markdown("### 📊 Direct ABC Observation Summary")
  abc_df = current_data["abc_data"]
  st.write(f"**Total Direct Records Ingested:** {len(abc_df)} logs")

  # Function summary table
  func_summary = (
      abc_df["Engine Auto-Inferred Function"]
      .value_counts()
      .reset_index(name="Count")
  )
  func_summary.columns = ["Inferred Function", "Count"]
  st.dataframe(func_summary, use_container_width=True, hide_index=True)

  st.markdown("**Clinical Trend Analysis:**")
  st.write(f"• **Primary Triggers:** {current_data['antecedents']}")
  st.write(f"• **Setting Events:** {current_data['setting_events']}")

st.divider()

# ==========================================
# 7. One-Click Structured FBA & BIP Report Preview
# ==========================================
st.markdown("## 📄 Legally & Clinically Aligned Report Preview")

with st.expander("👁️ View Full Synthesized FBA & BIP Document Structure", expanded=True):
  st.markdown(f"### FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) & BIP REPORT")
  st.markdown(f"**Client Name:** {current_data['client_meta']['name']} | **Age Cohort:** {current_data['client_meta']['age']}")
  st.markdown(f"**Agency / Placement:** {current_data['client_meta']['agency']}")
  
  st.markdown("---")
  st.markdown("#### 1. Target Behavior Operational Breakdown")
  st.write(f"**Description:** {current_data['target_behavior']}")
  st.write(f"**Examples:** {current_data['examples']}")
  st.write(f"**Non-Examples:** {current_data['non_examples']}")
  
  st.markdown("#### 2. Representative ABC Exemplars (Exemplar Subset)")
  st.table(abc_df[["Antecedent (A)", "Behavior (B)", "Consequence (C)"]].head(3))
  
  st.markdown("#### 3. Tailored Behavior Intervention Plan (BIP)")
  bip = current_data["bip_strategies"]
  st.markdown(f"**Antecedent Modifications:**\n{bip['antecedent']}")
  st.markdown(f"**Replacement Behaviors (FCT):**\n{bip['replacement']}")
  st.markdown(f"**Reinforcement Protocol:**\n{bip['reinforcement']}")
  st.markdown(f"**Reactive & Reduction Protocol:**\n{bip['reduction']}")
  st.markdown(f"**Crisis & Safety Plan:**\n{bip['crisis']}")

# ==========================================
# 8. Export Functionality (Word .docx with Appendix)
# ==========================================
def build_word_report(data):
  doc = docx.Document()
  for s in doc.sections:
    s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(
        0.7
    )

  # Title
  title_p = doc.add_heading(
      f"CLINICAL FBA & BIP REPORT\n[{data['framework']}]", level=0
  )
  title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

  # Meta Table
  info_table = doc.add_table(rows=3, cols=4)
  info_table.style = "Table Grid"
  meta = data["client_meta"]
  table_data = [
      [("Client Name", meta["name"]), ("Age Cohort", meta["age"])],
      [("DOB", meta["dob"]), ("Client ID", meta["id"])],
      [("Agency / Setting", meta["agency"]), ("Framework", data["framework"])],
  ]
  for r_idx, row in enumerate(table_data):
    for c_group, (lbl, val) in enumerate(row):
      cell_lbl = info_table.cell(r_idx, c_group * 2)
      cell_val = info_table.cell(r_idx, c_group * 2 + 1)
      cell_lbl.text = lbl
      cell_val.text = str(val)
      shd = parse_xml(r'<w:shd {} w:fill="F2F2F2"/>'.format(nsdecls("w")))
      cell_lbl._tc.get_or_add_tcPr().append(shd)
      cell_lbl.paragraphs[0].runs[0].font.bold = True
      cell_lbl.paragraphs[0].runs[0].font.size = Pt(8.5)
      cell_val.paragraphs[0].runs[0].font.size = Pt(8.5)

  doc.add_paragraph()

  # Section 1: Target Behavior
  doc.add_heading("1. Target Behavior Operational Breakdown", level=1)
  doc.add_paragraph(f"Description: {data['target_behavior']}")
  doc.add_paragraph(f"Examples: {data['examples']}")
  doc.add_paragraph(f"Non-Examples: {data['non_examples']}")

  # Section 2: Summary ABC & Trend Analysis
  doc.add_heading("2. Systematic ABC Trend Analysis & Summary", level=1)
  doc.add_paragraph(f"Primary Antecedent Triggers: {data['antecedents']}")
  doc.add_paragraph(f"Setting Events: {data['setting_events']}")

  # Section 3: BIP Interventions
  doc.add_heading("3. Behavior Intervention Plan (BIP)", level=1)
  bip = data["bip_strategies"]
  doc.add_heading("3.1 Antecedent Modifications", level=2)
  doc.add_paragraph(bip["antecedent"])
  doc.add_heading("3.2 Replacement Behaviors & Teaching Protocol", level=2)
  doc.add_paragraph(bip["replacement"])
  doc.add_heading("3.3 Reinforcement Schedule", level=2)
  doc.add_paragraph(bip["reinforcement"])
  doc.add_heading("3.4 Response Strategies", level=2)
  doc.add_paragraph(bip["reduction"])
  doc.add_heading("3.5 Crisis & De-escalation Protocol", level=2)
  doc.add_paragraph(bip["crisis"])

  # Appendix A: Full Raw ABC Data
  doc.add_page_break()
  doc.add_heading(
      f"Appendix A: Full Raw Direct ABC Observation Data ({len(data['abc_data'])} Logs)",
      level=1,
  )
  doc.add_paragraph(
      "The table below contains the complete unedited observation ledger"
      " appended for audit compliance."
  )

  raw_df = data["abc_data"]
  headers = list(raw_df.columns)
  table = doc.add_table(rows=1, cols=len(headers))
  table.style = "Table Grid"

  for idx, text in enumerate(headers):
    cell = table.rows[0].cells[idx]
    cell.text = text
    shd = parse_xml(r'<w:shd {} w:fill="1F4E78"/>'.format(nsdecls("w")))
    cell._tc.get_or_add_tcPr().append(shd)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in p.runs:
      r.font.bold = True
      r.font.color.rgb = RGBColor(255, 255, 255)
      r.font.size = Pt(8)

  for r_idx, row in raw_df.iterrows():
    row_cells = table.add_row().cells
    for c_idx, val in enumerate(row):
      row_cells[c_idx].text = str(val)
      row_cells[c_idx].paragraphs[0].runs[0].font.size = Pt(7.5)

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# Export Action Area
st.markdown("### 🚀 Export Production-Ready Clinical Artifacts")
btn_col1, btn_col2 = st.columns(2)

with btn_col1:
  doc_bytes = build_word_report(current_data)
  st.download_button(
      label=(
          "⬇️ Download Complete Aligned FBA & BIP Report (.docx) [With"
          " Appendix A]"
      ),
      data=doc_bytes,
      file_name=(
          f"Clinical_FBA_BIP_{st.session_state.active_group.upper()}.docx"
      ),
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      use_container_width=True,
      type="primary",
  )

with btn_col2:
  st.caption("ℹ️ Generates a fully formatted Word document containing the lean summary report and full Raw ABC Data Appendix.")

st.divider()
st.caption(
    "⚠️ **Clinical Decision Support Notice:** This system is designed as a decision-support sandbox for credentialed BCBAs, LBAs, and Special Education Directors. All generated intervention plans must be reviewed and signed off by a licensed professional prior to clinical implementation."
)
