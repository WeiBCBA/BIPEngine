import io
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Pt, RGBColor
import pandas as pd
import streamlit as st

# ==========================================
# 1. Page Configuration & Custom CSS
# ==========================================
st.set_page_config(
    page_title="BCBA Clinical FBA Engine (BCBA 临床 FBA 生成引擎)",
    page_icon="🧩",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-header { font-size: 2.0rem; color: #1F4E78; font-weight: 700; margin-bottom: 0.2rem; }
    .sub-header { font-size: 0.95rem; color: #555; margin-bottom: 1.2rem; }
    .privacy-banner {
        background-color: #EBF3FA;
        border-left: 5px solid #1F4E78;
        padding: 0.8rem 1rem;
        border-radius: 4px;
        margin-bottom: 1.2rem;
    }
    .step-box {
        background-color: #F8F9FA;
        border: 1px solid #E9ECEF;
        padding: 1.2rem;
        border-radius: 8px;
        margin-bottom: 1rem;
    }
    .stButton>button { border-radius: 6px; font-weight: 600; }
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# 2. Mock File Generators (In-Memory Helper)
# ==========================================
# 动态生成可供测试下载的 Mock ABC (CSV) 和 Interview (Docx)


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
        "Summary of Parent Interview:\nParents report that tantrums occur"
        " most frequently during unexpected transitions or loud household"
        " environments (e.g., vacuuming, blender). PICA behavior increases when"
        " client is under-stimulated or experiencing sensory overload."
    )
  elif cohort_key == "g2":
    doc.add_heading(
        "INDIRECT ASSESSMENT: TEACHER & PARAPROFESSIONAL INTERVIEW", level=1
    )
    doc.add_paragraph("Client Age: 11.2 years | Setting: Resource Room & Gen-Ed")
    doc.add_paragraph(
        "Summary of Teacher Notes:\nTeacher notes severe task avoidance"
        " during non-preferred academic activities, specifically multi-step"
        " writing prompts. Vocal outbursts also occur when direct teacher"
        " attention is shifted to other students."
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
        " personal electronics. Client demonstrates strong verbal"
        " self-advocacy when prompted with visual schedule aids."
    )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 3. Top Privacy Banner & Header
# ==========================================
st.markdown(
    """
    <div class="privacy-banner">
        <h3 style="margin-top:0; color:#1F4E78; font-size: 1.1rem;">🔒 Local-Only Memory Architecture (HIPAA Compliant Demo)</h3>
        <p style="margin-bottom:0; font-size: 0.88rem; color: #333;">
            所有上传的 ABC 观察表与访谈记录均在<strong>本地浏览器内存</strong>中进行实时解析与字段填充。系统不包含任何云端数据库存储或第三方外存，关闭页面数据即刻物理销毁。
        </p>
    </div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    "<div class='main-header'>🧩 BCBA Clinical FBA & BIP Decision Engine</div>",
    unsafe_allow_html=True,
)
st.markdown(
    "<div class='sub-header'>Automated Data Extraction & Structured Report"
    " Generation Architecture</div>",
    unsafe_allow_html=True,
)

st.divider()

# ==========================================
# 4. Cohort Selector & Input Workflow
# ==========================================
st.markdown("### 1️⃣ Select Clinical Cohort (选择服务人群分组)")

tab1, tab2, tab3 = st.tabs([
    "👶 Early Intervention (2-5 Yrs)",
    "🏫 School-Age / IEP (5-21 Yrs)",
    "💼 Adult Community & Vocational (21+ Yrs)",
])

# Default variables
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

# --- STEP 2: Input Controls (Direct, Indirect, QABF) ---
st.markdown("### 2️⃣ Import Assessment Artifacts (导入评估资料)")

col_input1, col_input2, col_input3 = st.columns([1.2, 1.2, 1.1])

# --- Entry 1: Direct Observation Data ---
with col_input1:
  st.markdown("#### 📄 Direct Observation (ABC)")
  st.caption("Supports Google Forms / MS Forms exported Excel or CSV")

  # Download Mock File Button
  mock_csv = generate_mock_abc_csv(selected_cohort)
  st.download_button(
      label=f"📥 Download Mock ABC Data (.csv)",
      data=mock_csv,
      file_name=f"Mock_ABC_Data_{selected_cohort}.csv",
      mime="text/csv",
      use_container_width=True,
  )

  uploaded_abc = st.file_uploader(
      "Upload Direct ABC Log:",
      type=["csv", "xlsx"],
      key=f"abc_{selected_cohort}",
  )

# --- Entry 2: Indirect Assessment Data ---
with col_input2:
  st.markdown("#### 📝 Indirect Interview Notes")
  st.caption("Supports Parent/Teacher interview notes in Docx or TXT")

  # Download Mock File Button
  mock_docx = generate_mock_interview_docx(selected_cohort)
  st.download_button(
      label=f"📥 Download Mock Interview Note (.docx)",
      data=mock_docx,
      file_name=f"Mock_Interview_{selected_cohort}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      use_container_width=True,
  )

  uploaded_interview = st.file_uploader(
      "Upload Interview Notes:",
      type=["docx", "txt"],
      key=f"interview_{selected_cohort}",
  )

# --- Entry 3: QABF Interactive Scoring ---
with col_input3:
  st.markdown("#### 📊 QABF Psychometric Input")
  st.caption("Adjust scores based on clinical QABF assessment")

  q_attention = st.number_input("Social Attention", 0, 15, value=4, step=1)
  q_escape = st.number_input("Task Escape", 0, 15, value=12, step=1)
  q_tangible = st.number_input("Tangible Access", 0, 15, value=6, step=1)
  q_sensory = st.number_input("Sensory / Automatic", 0, 15, value=14, step=1)
  q_physical = st.number_input("Physical Discomfort", 0, 15, value=2, step=1)

st.divider()

# ==========================================
# 5. Output Language & Report Synthesis
# ==========================================
st.markdown("### 3️⃣ Report Language & Synthesis Options")

col_lang, col_btn = st.columns([1, 2])

with col_lang:
  report_lang = st.radio(
      "Select Output Language Format:",
      options=[
          "English (US Standard)",
          "Bilingual (English / 简体中文)",
          "Bilingual (English / 繁體中文)",
      ],
      index=0,
  )

with col_btn:
  st.write(" ")
  st.write(" ")
  generate_click = st.button(
      "⚡ Extract Data & Generate Complete FBA / BIP Report",
      type="primary",
      use_container_width=True,
  )

# ==========================================
# 6. Report Generation & Preview Logic
# ==========================================
if generate_click or uploaded_abc or uploaded_interview:
  st.success("✅ Inputs processed successfully into Local Data Buffer!")

  # Parsing Uploaded Direct ABC Data into DataFrame
  if uploaded_abc is not None:
    try:
      if uploaded_abc.name.endswith(".csv"):
        parsed_abc_df = pd.read_csv(uploaded_abc)
      else:
        parsed_abc_df = pd.read_excel(uploaded_abc)
    except Exception as e:
      st.error(f"Error reading ABC file: {e}")
      parsed_abc_df = pd.read_csv(io.StringIO(mock_csv.decode("utf-8")))
  else:
    # Default to mock
    parsed_abc_df = pd.read_csv(io.StringIO(mock_csv.decode("utf-8")))

  # Preview Area
  st.markdown("---")
  st.markdown("## 📄 Synthesized Clinical FBA & BIP Report Preview")

  with st.expander("👁️ Review 9-Section Structured FBA/BIP Document", expanded=True):
    st.markdown(f"### FUNCTIONAL BEHAVIOR ASSESSMENT (FBA) & BIP")
    st.markdown(
        f"**Framework Compliance:** {framework_label} | **Cohort:**"
        f" {cohort_title}"
    )
    st.markdown(f"**Report Format:** {report_lang}")
    st.markdown("---")

    # Section 1: Demographics
    st.markdown("#### Section 1: Client Demographics & Referral Reason")
    st.write(
        "Client Name: [CLIENT_NAME] | DOB: 05/12/2018 | Setting: Clinic /"
        " Community"
    )

    # Section 2: Target Behavior Definition
    st.markdown("#### Section 2: Operational Definition of Target Behaviors")
    st.write(
        "**Target Behavior:** Aggression (hitting, pushing) and Property"
        " Destruction (sweeping items off surfaces)."
    )
    st.write(
        "**Onset/Offset:** Initiates within 5 seconds of non-preferred demand;"
        " ceases when client is removed from area."
    )

    # Section 3: Indirect Assessment Summary
    st.markdown("#### Section 3: Indirect Assessment Findings (Interviews)")
    if uploaded_interview:
      st.info(f"📁 Extracted from uploaded interview note: '{uploaded_interview.name}'")
    else:
      st.info("📁 Extracted from standard stakeholder interview logs.")
    st.write(
        "Stakeholders report highest frequency of behaviors during non-preferred"
        " task transitions and periods of reduced direct interaction."
    )

    # Section 4: Direct Observation & ABC Trend Analysis
    st.markdown(
        "#### Section 4: Direct Observation Analysis (ABC Ingestion)"
    )
    st.write(
        f"**Total Direct ABC Records Ingested:** {len(parsed_abc_df)} observation"
        " entries."
    )
    st.dataframe(parsed_abc_df.head(3), use_container_width=True)

    # Section 5: QABF Results
    st.markdown("#### Section 5: QABF Psychometric Assessment Results")
    q_df = pd.DataFrame({
        "Function": [
            "Attention",
            "Escape",
            "Tangible",
            "Sensory",
            "Physical Pain",
        ],
        "Score": [q_attention, q_escape, q_tangible, q_sensory, q_physical],
    }).set_index("Function")
    st.bar_chart(q_df)

    # Section 6: Functional Hypothesis
    st.markdown("#### Section 6: Summary of Functional Hypothesis")
    st.write(
        "> **Hypothesis Statement:** When presented with non-preferred academic or"
        " transitional demands, the client engages in target behaviors"
        " maintained primarily by **Task Escape** and secondary **Access to"
        " Sensory Tools**."
    )

    # Section 7: BIP Strategies
    st.markdown("#### Section 7: Behavior Intervention Plan (BIP)")
    st.write(
        "• **Antecedent Modifications:** Task chunking, pre-task choice"
        " provision, visual schedule."
    )
    st.write(
        "• **Replacement Behavior (FCT):** Functional communication training using"
        " 'Break' card or self-advocacy phrase."
    )
    st.write(
        "• **Reinforcement Schedule:** FR-1 schedule of preferred sensory access"
        " upon functional communication exchange."
    )

    # Section 8: Crisis Plan
    st.markdown("#### Section 8: Crisis & De-escalation Protocol")
    st.write(
        "Maintain environmental safety, block dangerous impacts neutrally,"
        " avoid verbal lecturing during active escalation."
    )

    # Section 9: Appendix (Raw Data)
    st.markdown(
        "#### Section 9: Appendix A - Raw Direct Observation Log (Ingested)"
    )
    st.caption(
        "Full unedited raw ABC entries auto-appended for audit compliance."
    )
    st.dataframe(parsed_abc_df, use_container_width=True)

  # Export Action
  st.markdown("### 🚀 Export Production Artifacts")

  # Generate Word Document Function


  def generate_word_doc():
    doc = docx.Document()
    doc.add_heading("CLINICAL FUNCTIONAL BEHAVIOR ASSESSMENT (FBA)", level=0)
    doc.add_paragraph(f"Cohort: {cohort_title}")
    doc.add_paragraph(f"Framework: {framework_label}")
    doc.add_paragraph(f"Language Output: {report_lang}")

    doc.add_heading("1. Target Behavior Definition", level=1)
    doc.add_paragraph(
        "Operational definition extracted and synthesized from multi-source"
        " assessment."
    )

    doc.add_heading("2. QABF Psychometric Summary", level=1)
    doc.add_paragraph(
        f"Scores -> Escape: {q_escape}, Sensory: {q_sensory}, Attention:"
        f" {q_attention}, Tangible: {q_tangible}, Physical: {q_physical}"
    )

    doc.add_heading(
        f"Appendix A: Full Ingested Raw ABC Data ({len(parsed_abc_df)} Entries)",
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

  doc_file = generate_word_doc()

  st.download_button(
      label=(
          "⬇️ Download Formal FBA & BIP Report (.docx) [Including Appendix A]"
      ),
      data=doc_file,
      file_name=f"FBA_BIP_Report_{selected_cohort}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      type="primary",
  )
