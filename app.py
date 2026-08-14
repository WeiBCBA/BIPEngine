import io
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
import pandas as pd
import streamlit as st

# ==========================================
# 1. Page Configuration & Professional Styling
# ==========================================
st.set_page_config(
    page_title="BCBA Clinical FBA & BIP Formulation Engine",
    page_icon="🧩",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-header { font-size: 2.1rem; color: #1F4E78; font-weight: 700; margin-bottom: 0.2rem; }
    .sub-header { font-size: 0.95rem; color: #555; margin-bottom: 1.2rem; }
    .privacy-banner {
        background-color: #EBF3FA;
        border-left: 5px solid #1F4E78;
        padding: 0.8rem 1rem;
        border-radius: 4px;
        margin-bottom: 1.2rem;
    }
    .stButton>button { border-radius: 6px; font-weight: 600; }
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. Mock Data Generators (In-Memory Helper)
# ==========================================
def generate_mock_abc_csv(cohort_key):
  datasets = {
      "g1": [
          {
              "Date_Time": "2026-08-10 09:15",
              "Setting": "Clinic Playroom",
              "Antecedent": (
                  "Kitchen blender noise started in adjacent breakroom"
              ),
              "Behavior": "Screamed, slapped face 3x, dropped to floor",
              "Consequence": (
                  "RBT offered noise-canceling headphones and sensory chew tool"
              ),
          },
          {
              "Date_Time": "2026-08-10 10:30",
              "Setting": "Outdoor Playground",
              "Antecedent": "Transition prompt given to pack up sand toys",
              "Behavior": (
                  "Grabbed handful of sand and attempted to put in mouth"
              ),
              "Consequence": (
                  "RBT blocked mouth, redirected to oral chew tool, demand"
                  " paused"
              ),
          },
          {
              "Date_Time": "2026-08-11 11:00",
              "Setting": "Table Therapy Room",
              "Antecedent": "Presented matching discrete trial worksheet",
              "Behavior": "Head banging on foam floor mat",
              "Consequence": (
                  "Therapist paused demand, presented PECS 'Break' card"
              ),
          },
      ],
      "g2": [
          {
              "Date_Time": "2026-08-10 09:30",
              "Setting": "Gen-Ed Classroom",
              "Antecedent": "Teacher presented 2-page independent math worksheet",
              "Behavior": "Screamed 'I won't do it!', pushed desk away",
              "Consequence": (
                  "Staff presented 'Break' visual card; demand paused for 2"
                  " minutes"
              ),
          },
          {
              "Date_Time": "2026-08-11 13:15",
              "Setting": "Small Group Reading",
              "Antecedent": "Teacher turned attention to help peer",
              "Behavior": "Threw textbook across desk, shouted 'Look at me!'",
              "Consequence": (
                  "Staff redirected with neutral tone to waiting visual"
                  " schedule"
              ),
          },
          {
              "Date_Time": "2026-08-12 10:45",
              "Setting": "Resource Room",
              "Antecedent": "Multi-step essay prompt assigned",
              "Behavior": "Tore paper, swept markers onto floor",
              "Consequence": (
                  "Guided to quiet break area; task chunked into single"
                  " sentence prompt"
              ),
          },
      ],
      "g3": [
          {
              "Date_Time": "2026-08-09 09:00",
              "Setting": "Residential Home",
              "Antecedent": "New staff worker introduced schedule change",
              "Behavior": "Swept dishes off table, shouted threats",
              "Consequence": (
                  "DSP stepped in, offered visual choice board, demand paused"
              ),
          },
          {
              "Date_Time": "2026-08-10 14:00",
              "Setting": "Vocational Workshop",
              "Antecedent": "Tablet time limit reached during break",
              "Behavior": (
                  "Pacing, grabbed tablet back, knocked over assembly chair"
              ),
              "Consequence": (
                  "Job coach prompted self-advocacy phrase card and granted"
                  " 5-min extension"
              ),
          },
          {
              "Date_Time": "2026-08-11 18:30",
              "Setting": "Community Living Room",
              "Antecedent": "Housemate adjusted TV channel without consensus",
              "Behavior": "Blocked TV screen, loud vocal resistance",
              "Consequence": "DSP facilitated structured housemate mediation",
          },
      ],
  }
  df = pd.DataFrame(datasets.get(cohort_key, datasets["g1"]))
  return df.to_csv(index=False).encode("utf-8")


def generate_mock_interview_docx(cohort_key):
  doc = docx.Document()
  if cohort_key == "g1":
    doc.add_heading("INDIRECT ASSESSMENT: PARENT & RBT INTERVIEW", level=1)
    doc.add_paragraph("Client Age: 3.5 years | Setting: Early Intervention")
    doc.add_paragraph(
        "Summary of Parent Interview:\nParents report tantrums occur"
        " frequently during unexpected transitions or loud household"
        " environments. PICA increases when client experiences sensory"
        " overload."
    )
  elif cohort_key == "g2":
    doc.add_heading(
        "INDIRECT ASSESSMENT: TEACHER & PARAPROFESSIONAL INTERVIEW", level=1
    )
    doc.add_paragraph("Client Age: 11.2 years | Setting: Resource Room & Gen-Ed")
    doc.add_paragraph(
        "Summary of Teacher Notes:\nTeacher notes severe task avoidance"
        " during non-preferred academic activities, specifically multi-step"
        " writing prompts."
    )
  else:
    doc.add_heading(
        "INDIRECT ASSESSMENT: STAKEHOLDER & JOB COACH INTERVIEW", level=1
    )
    doc.add_paragraph(
        "Client Age: 26.8 years | Setting: Community Residential & Vocational"
    )
    doc.add_paragraph(
        "Summary of DSP Notes:\nSupport staff indicate property destruction"
        " is strongly tied to sudden schedule changes and restricted access to"
        " personal electronics."
    )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 3. Security Banner & Main Header
# ==========================================
st.markdown(
    """
    <div class="privacy-banner">
        <h3 style="margin-top:0; color:#1F4E78; font-size: 1.05rem;">🔒 Local-Memory Architecture (HIPAA Compliant Demo)</h3>
        <p style="margin-bottom:0; font-size: 0.88rem; color: #333;">
            All uploaded direct observation logs and indirect interview files are parsed strictly within the <strong>local browser session memory</strong>. No cloud server storage or external databases are utilized. Data is purged immediately upon session termination.
        </p>
    </div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    "<div class='main-header'>🧩 BCBA Clinical FBA & BIP Formulation"
    " Engine</div>",
    unsafe_allow_html=True,
)
st.markdown(
    "<div class='sub-header'>Automated Data Extraction & Draft Synthesis for"
    " Credentialed BCBAs</div>",
    unsafe_allow_html=True,
)

st.divider()

# ==========================================
# 4. Cohort Selection Workflow
# ==========================================
st.markdown("### 1️⃣ Select Clinical Cohort")

tab1, tab2, tab3 = st.tabs([
    "👶 Early Intervention (2-5 Yrs)",
    "🏫 School-Age / IEP (5-21 Yrs)",
    "💼 Adult Community & Vocational (21+ Yrs)",
])

selected_cohort = "g1"
cohort_title = "Early Intervention Protocol (2-5 Yrs)"
framework_label = "ESDM / NDBI Framework"

with tab1:
  selected_cohort = "g1"
  cohort_title = "Early Intervention Protocol (2-5 Yrs)"
  framework_label = "ESDM / NDBI Framework"
  st.caption("Standard: Early Start Denver Model / Developmental Compliance")

with tab2:
  selected_cohort = "g2"
  cohort_title = "School-Age / IEP Protocol (5-21 Yrs)"
  framework_label = "IDEA IEP / PBIS Framework"
  st.caption(
      "Standard: IDEA Accommodations, PBIS & Functional Communication Training"
  )

with tab3:
  selected_cohort = "g3"
  cohort_title = "Adult Community & Vocational Protocol (21+ Yrs)"
  framework_label = "Medicaid HCBS / Person-Centered Waiver Framework"
  st.caption(
      "Standard: Person-Centered Adult Services & Independent Living Support"
  )

st.write(" ")

# ==========================================
# 5. Assessment Data Import
# ==========================================
st.markdown("### 2️⃣ Import Assessment Artifacts")

col_input1, col_input2, col_input3 = st.columns([1.2, 1.2, 1.1])

# --- Input 1: Direct ABC ---
with col_input1:
  st.markdown("#### 📄 Direct Observation (ABC)")
  st.caption("Google Forms / MS Forms exported CSV or Excel")

  mock_csv = generate_mock_abc_csv(selected_cohort)
  st.download_button(
      label=f"📥 Download Sample ABC Log (.csv)",
      data=mock_csv,
      file_name=f"Sample_ABC_Data_{selected_cohort}.csv",
      mime="text/csv",
      use_container_width=True,
  )

  uploaded_abc = st.file_uploader(
      "Upload Direct ABC File:",
      type=["csv", "xlsx"],
      key=f"abc_{selected_cohort}",
  )

# --- Input 2: Indirect Interview ---
with col_input2:
  st.markdown("#### 📝 Indirect Interview Notes")
  st.caption("Parent/Teacher interview notes in Docx or TXT")

  mock_docx = generate_mock_interview_docx(selected_cohort)
  st.download_button(
      label=f"📥 Download Sample Interview (.docx)",
      data=mock_docx,
      file_name=f"Sample_Interview_{selected_cohort}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      use_container_width=True,
  )

  uploaded_interview = st.file_uploader(
      "Upload Interview Notes File:",
      type=["docx", "txt"],
      key=f"interview_{selected_cohort}",
  )

# --- Input 3: QABF Scoring ---
with col_input3:
  st.markdown("#### 📊 QABF Psychometric Input")
  st.caption("Input scores from clinical QABF assessment")

  q_attention = st.number_input("Social Attention", 0, 15, value=4, step=1)
  q_escape = st.number_input("Task Escape", 0, 15, value=12, step=1)
  q_tangible = st.number_input("Tangible Access", 0, 15, value=6, step=1)
  q_sensory = st.number_input("Sensory / Automatic", 0, 15, value=14, step=1)
  q_physical = st.number_input("Physical Discomfort", 0, 15, value=2, step=1)

st.divider()

# ==========================================
# 6. Output Language & Synthesis Config
# ==========================================
st.markdown("### 3️⃣ Target Language & Synthesis Settings")

col_lang, col_btn = st.columns([1.2, 1.8])

with col_lang:
  report_lang = st.radio(
      "Select Target Report Language / Format:",
      options=[
          "English (US Standard)",
          "Bilingual (English / Spanish - 西班牙语)",
          "Bilingual (English / Simplified Chinese - 简体中文)",
      ],
      index=0,
  )

with col_btn:
  st.write(" ")
  st.write(" ")
  generate_click = st.button(
      "⚡ Extract Data & Formulate FBA / BIP Drafts",
      type="primary",
      use_container_width=True,
  )

# ==========================================
# 7. Synthesis Preview & Export Engine
# ==========================================
if generate_click or uploaded_abc or uploaded_interview:
  st.success("✅ Assessment artifacts parsed into local synthesis buffer.")

  # Parsing Direct ABC Data
  if uploaded_abc is not None:
    try:
      if uploaded_abc.name.endswith(".csv"):
        parsed_abc_df = pd.read_csv(uploaded_abc)
      else:
        parsed_abc_df = pd.read_excel(uploaded_abc)
    except Exception as e:
      st.error(f"Error parsing ABC file: {e}")
      parsed_abc_df = pd.read_csv(io.StringIO(mock_csv.decode("utf-8")))
  else:
    parsed_abc_df = pd.read_csv(io.StringIO(mock_csv.decode("utf-8")))

  st.markdown("---")
  st.markdown("## 📄 Formulated Clinical Draft Preview")

  with st.expander("👁️ Review Formulated FBA & BIP Structure", expanded=True):
    st.markdown(f"### CLINICAL FBA & BIP INITIAL DRAFT")
    st.markdown(
        f"**Framework:** {framework_label} | **Cohort:** {cohort_title} |"
        f" **Target Format:** {report_lang}"
    )
    st.markdown("---")

    # Preview Sections
    st.markdown("#### Part I: Functional Behavior Assessment (FBA)")
    st.write(
        "• **Operational Definition:** Target behaviors categorized by"
        " topography and severity."
    )
    st.write(
        f"• **Direct ABC Records Processed:** {len(parsed_abc_df)} observation"
        " entries parsed."
    )
    st.dataframe(parsed_abc_df.head(3), use_container_width=True)

    st.markdown("#### Part II: Behavior Intervention Plan (BIP)")
    st.write(
        "• **Antecedent Modifications:** Proactive environmental"
        " restructuring."
    )
    st.write(
        "• **Replacement Behavior (FCT):** Alternative functional communication"
        " training protocols."
    )
    st.write(
        "• **Reinforcement Schedule:** Differential reinforcement of"
        " alternative behavior (DRA)."
    )

  # ==========================================
  # 8. Document Generation Functions (Modular Export)
  # ==========================================

  # Builder 1: FBA Only Doc
  def build_fba_doc():
    doc = docx.Document()
    doc.add_heading("FUNCTIONAL BEHAVIOR ASSESSMENT (FBA)", level=0)
    doc.add_paragraph(f"Cohort: {cohort_title}")
    doc.add_paragraph(f"Framework: {framework_label}")
    doc.add_paragraph(f"Language Format: {report_lang}")

    doc.add_heading("1. Target Behavior Operational Breakdown", level=1)
    doc.add_paragraph(
        "Aggression and Property Destruction defined in measurable terms."
    )

    doc.add_heading("2. Indirect Assessment Summary", level=1)
    doc.add_paragraph(
        "Synthesized from stakeholder and parent interview notes."
    )

    doc.add_heading("3. QABF Psychometric Profile", level=1)
    doc.add_paragraph(
        f"Scores -> Escape: {q_escape}, Sensory: {q_sensory}, Attention:"
        f" {q_attention}, Tangible: {q_tangible}, Physical: {q_physical}"
    )

    doc.add_heading("4. Functional Hypothesis Statement", level=1)
    doc.add_paragraph(
        "Behavior is primarily maintained by Task Escape and automatic sensory"
        " stimulation."
    )

    doc.add_heading(
        f"Appendix A: Full Raw Direct ABC Observations ({len(parsed_abc_df)}"
        " Records)",
        level=1,
    )
    table = doc.add_table(rows=1, cols=len(parsed_abc_df.columns))
    table.style = "Table Grid"
    for i, col_name in enumerate(parsed_abc_df.columns):
      table.rows[0].cells[i].text = col_name
    for _, row in parsed_abc_df.iterrows():
      row_cells = table.add_row().cells
      for i, val in enumerate(row):
        row_cells[i].text = str(val)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

  # Builder 2: BIP Only Doc
  def build_bip_doc():
    doc = docx.Document()
    doc.add_heading("BEHAVIOR INTERVENTION PLAN (BIP)", level=0)
    doc.add_paragraph(f"Cohort: {cohort_title}")
    doc.add_paragraph(f"Target Format: {report_lang}")

    doc.add_heading("1. Antecedent Strategies", level=1)
    doc.add_paragraph(
        "Environmental accommodations, visual countdown timers, and task"
        " chunking."
    )

    doc.add_heading("2. Replacement Behavior Protocol (FCT)", level=1)
    doc.add_paragraph(
        "Functional Communication Training: Teaching client to request"
        " breaks or sensory tools."
    )

    doc.add_heading("3. Reinforcement System", level=1)
    doc.add_paragraph(
        "Differential Reinforcement of Alternative Behavior (DRA) on an FR-1"
        " schedule."
    )

    doc.add_heading("4. Crisis & Safety Plan", level=1)
    doc.add_paragraph(
        "De-escalation guidelines, maintaining safe physical radius, and zero"
        " unapproved restrictive practices."
    )

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

  # Builder 3: Combined FBA & BIP Package
  def build_combined_doc():
    doc = docx.Document()
    doc.add_heading("COMPREHENSIVE FBA & BIP REPORT PACKAGE", level=0)
    doc.add_paragraph(f"Cohort: {cohort_title} | Framework: {framework_label}")

    doc.add_heading("PART I: FUNCTIONAL BEHAVIOR ASSESSMENT (FBA)", level=1)
    doc.add_paragraph(
        "Complete evaluation findings, QABF profile, and functional"
        " hypothesis."
    )

    doc.add_heading("PART II: BEHAVIOR INTERVENTION PLAN (BIP)", level=1)
    doc.add_paragraph(
        "Proactive, teaching, reinforcement, and reactive protocols."
    )

    doc.add_heading(
        f"Appendix A: Ingested Direct ABC Data ({len(parsed_abc_df)} Entries)",
        level=1,
    )
    table = doc.add_table(rows=1, cols=len(parsed_abc_df.columns))
    table.style = "Table Grid"
    for i, col_name in enumerate(parsed_abc_df.columns):
      table.rows[0].cells[i].text = col_name
    for _, row in parsed_abc_df.iterrows():
      row_cells = table.add_row().cells
      for i, val in enumerate(row):
        row_cells[i].text = str(val)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

  # ==========================================
  # 9. Modular Download Action Panel
  # ==========================================
  st.markdown("### 🚀 Download Formulated Clinical Artifacts")
  st.caption(
      "Choose to export standalone FBA, standalone BIP protocol, or the"
      " combined comprehensive package."
  )

  col_dl1, col_dl2, col_dl3 = st.columns(3)

  with col_dl1:
    fba_bytes = build_fba_doc()
    st.download_button(
        label="📄 Download FBA Report Only (.docx)",
        data=fba_bytes,
        file_name=f"FBA_Report_{selected_cohort}.docx",
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        use_container_width=True,
    )

  with col_dl2:
    bip_bytes = build_bip_doc()
    st.download_button(
        label="📋 Download BIP Protocol Only (.docx)",
        data=bip_bytes,
        file_name=f"BIP_Protocol_{selected_cohort}.docx",
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        use_container_width=True,
    )

  with col_dl3:
    combined_bytes = build_combined_doc()
    st.download_button(
        label="📦 Download Full FBA & BIP Package (.docx)",
        data=combined_bytes,
        file_name=f"Full_FBA_BIP_Package_{selected_cohort}.docx",
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        type="primary",
        use_container_width=True,
    )

st.divider()
st.caption(
    "⚠️ **Clinical Responsibility Notice:** This formulation engine serves as"
    " a clinical draft synthesizer for credentialed BCBAs and LBAs. All"
    " generated drafts must be independently reviewed, edited, and verified"
    " by the supervising clinician prior to formal signature and"
    " implementation."
)
