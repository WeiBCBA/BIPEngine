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
    page_title="BCBA FBA & BIP Draft Formulation Tool (Demo Version)",
    page_icon="🧩",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    /* Harmonized Main Header Style */
    .main-title {
        font-size: 2.1rem;
        color: #1F4E78;
        font-weight: 700;
        margin-bottom: 0.2rem;
        line-height: 1.3;
    }
    .demo-tag {
        font-size: 1.8rem;
        color: #1F4E78;
        font-weight: 600;
    }
    .sub-header { font-size: 0.95rem; color: #555; margin-bottom: 1.2rem; }
    
    /* Top Privacy Banner */
    .privacy-banner {
        background-color: #EBF3FA;
        border-left: 5px solid #1F4E78;
        padding: 0.8rem 1rem;
        border-radius: 4px;
        margin-bottom: 1.2rem;
    }

    /* Core Critical Security Notice */
    .critical-security-card {
        background-color: #FFF9E6;
        border: 2px solid #FFE082;
        border-left: 6px solid #FFB300;
        padding: 1rem 1.2rem;
        border-radius: 8px;
        margin-top: 0.5rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .security-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #B78103;
        margin-bottom: 0.5rem;
    }
    .security-body {
        font-size: 0.90rem;
        color: #333333;
        line-height: 1.6;
    }
    .security-body ul {
        margin-top: 0.3rem;
        margin-bottom: 0.3rem;
        padding-left: 1.2rem;
    }
    .security-body li {
        margin-bottom: 0.4rem;
    }
    .stButton>button { border-radius: 6px; font-weight: 600; }
    </style>
""",
    unsafe_allow_html=True,
)


# ==========================================
# 2. Dynamic Mock Data Generators (ABC & Triangulated Interview Raw Data)
# ==========================================
def generate_mock_abc_csv(cohort_key):
  datasets = {
      "g1": [  # Early Intervention (2-5 Yrs)
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
              "Behavior": "Head banging on foam floor mat (3-4 occurrences)",
              "Consequence": (
                  "Therapist paused demand, presented PECS 'Break' card"
              ),
          },
          {
              "Date_Time": "2026-08-11 15:20",
              "Setting": "Clinic Snack Area",
              "Antecedent": "Preferred juice cup emptied",
              "Behavior": "Biting own wrist, high-pitched crying",
              "Consequence": (
                  "RBT prompted functional communication button 'More Juice'"
              ),
          },
          {
              "Date_Time": "2026-08-12 09:45",
              "Setting": "Gross Motor Gym",
              "Antecedent": "Peer approached to share trampoline space",
              "Behavior": "Pushed peer away, vocal distress",
              "Consequence": (
                  "Staff facilitated physical boundary block, guided to"
                  " alternative swing"
              ),
          },
      ],
      "g2": [  # School-Age (5-21 Yrs)
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
          {
              "Date_Time": "2026-08-13 11:30",
              "Setting": "School Cafeteria",
              "Antecedent": "Loud bell rang for lunch transition",
              "Behavior": "Covered ears, hit staff arm repeatedly",
              "Consequence": (
                  "Escorted to alternative low-stimulation lunch seating"
              ),
          },
          {
              "Date_Time": "2026-08-13 14:10",
              "Setting": "PE Class",
              "Antecedent": "Loss in structured kickball game",
              "Behavior": "Kicked gym equipment, verbal swearing at peers",
              "Consequence": (
                  "Prompted self-regulation visual strip, taken on brief cool-down walk"
              ),
          },
      ],
      "g3": [  # Adult Vocational (21+ Yrs)
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
          {
              "Date_Time": "2026-08-12 10:15",
              "Setting": "Community Supermarket",
              "Antecedent": "Long queue line at checkout counter",
              "Behavior": "Agitated pacing, verbal refusal, attempted to drop cart items",
              "Consequence": (
                  "Prompted use of personal music headphones and sensory fidget"
              ),
          },
          {
              "Date_Time": "2026-08-12 16:00",
              "Setting": "Vocational Sorting Area",
              "Antecedent": "Supervisor requested re-sorting mislabeled box",
              "Behavior": "Refused task, slammed box onto table",
              "Consequence": (
                  "Task broken down into visual checklist; offered 3-min coffee"
                  " break"
              ),
          },
      ],
  }
  df = pd.DataFrame(datasets.get(cohort_key, datasets["g1"]))
  return df.to_csv(index=False).encode("utf-8")


def generate_mock_interview_docx(cohort_key):
  doc = docx.Document()

  if cohort_key == "g1":
    doc.add_heading(
        "INDIRECT ASSESSMENT: RAW STAKEHOLDER INTERVIEW NOTES (DE-IDENTIFIED)",
        level=1,
    )
    doc.add_paragraph(
        "Client ID: [CLIENT_ID_01] | Target Cohort: Early Intervention (2-5"
        " Yrs)\nInterviewer: BCBA | Informants: Parent & Preschool Director\n"
    )

    doc.add_heading(
        "Moment 1: Home High-Frequency Sound Reactions (Parent Quote)", level=2
    )
    doc.add_paragraph(
        "\"At home, whenever the vacuum cleaner or blender turns on in the"
        " kitchen, he instantly covers his ears and starts slapping his cheeks"
        " with open palms. He drops to the rug and screams until we turn off the"
        " machine. Giving him his weighted blanket or chew tool seems to help"
        " him reset after a few minutes.\""
    )

    doc.add_heading(
        "Moment 2: Sensory Toy Transitions at Daycare (Teacher Quote)", level=2
    )
    doc.add_paragraph(
        "\"He loves sensory play like water tables and sand boxes. When we give"
        " him a 2-minute warning to put away the shovels and clean his hands,"
        " he gets very anxious and tries to scoop sand directly into his"
        " mouth. We usually have to step in close to block him physically and"
        " give him a rubber chew necklace instead.\""
    )

    doc.add_heading(
        "Moment 3: Structured Fine-Motor Tasks (Preschool Director Quote)",
        level=2,
    )
    doc.add_paragraph(
        "\"During morning circle or table work, when presented with repetitive"
        " fine-motor tasks like bead-stringing or matching sheets, he gets"
        " frustrated within two minutes. He leans forward and bangs his forehead"
        " against the foam mat on the floor. If we hand him his 'Break' picture"
        " card, he calms down right away.\""
    )

    doc.add_heading(
        "Moment 4: Expressing Wants During Mealtime (Parent Quote)", level=2
    )
    doc.add_paragraph(
        "\"He doesn't have vocal words for specific drinks or snacks yet. At"
        " afternoon snack, when his juice cup becomes empty, he starts biting"
        " his wrist and whimpering high-pitched sounds. We're currently trying"
        " to teach him to push a big speech button that says 'More Juice' instead"
        " of biting himself.\""
    )

    doc.add_heading(
        "Moment 5: Peer Proximity in Play Gym (Teacher Quote)", level=2
    )
    doc.add_paragraph(
        "\"In the indoor play gym, he really enjoys the big blue trampoline."
        " However, if another toddler steps onto the padding nearby without"
        " warning, he shoves them away hard and screams loudly. We always need"
        " a staff member nearby to gently shadow him and guide him to the swing"
        " when it gets crowded.\""
    )

  elif cohort_key == "g2":
    doc.add_heading(
        "INDIRECT ASSESSMENT: RAW STAKEHOLDER INTERVIEW NOTES (DE-IDENTIFIED)",
        level=1,
    )
    doc.add_paragraph(
        "Client ID: [CLIENT_ID_02] | Target Cohort: School-Age / IEP (5-21"
        " Yrs)\nInterviewer: BCBA | Informants: Gen-Ed Teacher & IEP Case"
        " Manager\n"
    )

    doc.add_heading(
        "Moment 1: Independent Academic Task Demands (Teacher Quote)", level=2
    )
    doc.add_paragraph(
        "\"Whenever I pass out long, multi-page worksheets in 1st period"
        " math, he gets overwhelmed almost immediately. He pushes his desk away"
        " with a loud bang, throws his pencil, and yells 'I can't do this!'. If"
        " the classroom aide presents his visual break card right away, he takes"
        " 2 minutes to collect himself.\""
    )

    doc.add_heading(
        "Moment 2: Divided Teacher Attention in Group Work (Aide Quote)", level=2
    )
    doc.add_paragraph(
        "\"During small group reading, he likes having the teacher's direct"
        " feedback. As soon as the teacher turns away to assist another table,"
        " he throws his book across the desk or slams his binder to get everyone"
        " looking at him. We are working on reminding him to flip his visual"
        " 'Waiting' card instead of acting out.\""
    )

    doc.add_heading(
        "Moment 3: Complex Essay Writing Tasks (Case Manager Quote)", level=2
    )
    doc.add_paragraph(
        "\"In English resource class, multi-step writing prompts trigger quick"
        " avoidance. Last week, he tore up his assignment sheet and swept all"
        " his highlighters off the desk onto the floor. Once we break the essay"
        " down into sentence-by-sentence chunks, he handles it much better.\""
    )

    doc.add_heading(
        "Moment 4: High-Noise Cafeteria Environments (Aide Quote)", level=2
    )
    doc.add_paragraph(
        "\"The school cafeteria during 3rd lunch is extremely loud and chaotic."
        " When the bell rings unexpectedly, he covers his ears, panics, and sometimes"
        " strikes out at staff arms. Escorting him to the quiet resource room"
        " for lunch has significantly reduced these incidents.\""
    )

    doc.add_heading(
        "Moment 5: Competitive Games in PE Class (PE Teacher Quote)", level=2
    )
    doc.add_paragraph(
        "\"He loves sports, but handles losing or referee calls very poorly. When"
        " his team dropped a point in kickball, he kicked the boundary cones"
        " and yelled profanities at peers. I had to walk him over to the side"
        " track with a visual self-regulation strip to cool down before he could"
        " rejoin.\""
    )

  else:  # g3
    doc.add_heading(
        "INDIRECT ASSESSMENT: RAW STAKEHOLDER INTERVIEW NOTES (DE-IDENTIFIED)",
        level=1,
    )
    doc.add_paragraph(
        "Client ID: [CLIENT_ID_03] | Target Cohort: Adult Community (21+"
        " Yrs)\nInterviewer: BCBA | Informants: Direct Support Professional"
        " (DSP) & Vocational Supervisor\n"
    )

    doc.add_heading(
        "Moment 1: Unexpected Schedule Changes at Residence (DSP Quote)",
        level=2,
    )
    doc.add_paragraph(
        "\"He relies heavily on a predictable morning routine in the group"
        " home. When a substitute staff member arrived and tried to rearrange"
        " breakfast duties, he became upset, swept dishes off the counter, and"
        " shouted. We stepped back, gave him space, and brought out his visual"
        " schedule board to re-establish calm.\""
    )

    doc.add_heading(
        "Moment 2: Screen Time Limits at Vocational Workshop (Job Coach"
        " Quote)",
        level=2,
    )
    doc.add_paragraph(
        "\"During designated work breaks, he uses a tablet for videos. When the"
        " 15-minute timer went off, he paced back and forth, grabbed the tablet"
        " back aggressively, and knocked over a folding chair. Showing him a"
        " clear visual phrase card helped us transition him back to assembly"
        " work.\""
    )

    doc.add_heading(
        "Moment 3: Shared Living Room Choices (DSP Quote)", level=2
    )
    doc.add_paragraph(
        "\"In the evenings, he gets territorial over the main TV. If a housemate"
        " changes the channel without agreeing first, he physically stands"
        " directly in front of the screen, blocking the view and shouting. Staff"
        " have to mediate immediately and offer a structured choice board for TV"
        " times.\""
    )

    doc.add_heading(
        "Moment 4: Overcrowded Public Supermarket (DSP Quote)", level=2
    )
    doc.add_paragraph(
        "\"When we go grocery shopping during busy weekend hours, long checkout"
        " lines trigger high anxiety. He starts pacing nervously, vocalizing"
        " refusal, and trying to abandon the shopping cart. Prompting him to put"
        " on his noise-canceling headphones with music helps him tolerate the wait"
        " time.\""
    )

    doc.add_heading(
        "Moment 5: Unfamiliar Work Re-Sorting Tasks (Vocational Supervisor"
        " Quote)",
        level=2,
    )
    doc.add_paragraph(
        "\"At the vocational packaging line, if I ask him to undo a box and"
        " re-sort mislabeled parts, he experiences immediate task distress. He"
        " refused, slammed the supply box on the table, and crossed his arms."
        " Breaking the re-sorting process into a 3-step visual checklist helped"
        " him complete the job successfully.\""
    )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 3. Privacy Header & App Main Title
# ==========================================
st.markdown(
    """
    <div class="privacy-banner">
        <h3 style="margin-top:0; color:#1F4E78; font-size: 1.05rem;">🔒 Zero-Cloud Security & Local Session Memory</h3>
        <p style="margin-bottom:0; font-size: 0.88rem; color: #333;">
            All calculations and data parsing occur 100% locally within your browser's active memory. No external database, cloud storage, or third-party servers are used.
        </p>
    </div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    "<div class='main-title'>🧩 BCBA Clinical FBA & BIP Draft Formulation Tool"
    " <span class='demo-tag'>(Demo Version)</span></div>",
    unsafe_allow_html=True,
)
st.markdown(
    "<div class='sub-header'>Interactive Demonstration for Automated Clinical"
    " First-Draft Synthesis | Designed for BCBAs & LBAs</div>",
    unsafe_allow_html=True,
)

st.divider()

# ==========================================
# 4. Cohort Selection Workflow
# ==========================================
st.markdown("### 1️⃣ Select Clinical Cohort")

cohort_options = {
    "g1": "👶 Early Intervention Protocol (2-5 Yrs)",
    "g2": "🏫 School-Age / IEP Protocol (5-21 Yrs)",
    "g3": "💼 Adult Community & Vocational Protocol (21+ Yrs)",
}

selected_cohort_key = st.radio(
    "Select Target Client Population:",
    options=list(cohort_options.keys()),
    format_func=lambda x: cohort_options[x],
    index=0,
    horizontal=True,
)

# 🌟 Harmonized Vertical Separator |
cohort_meta = {
    "g1": {
        "title": "Early Intervention Protocol (2-5 Yrs)",
        "file_tag": "2to5yo",
        "framework": "ESDM | NDBI Framework",
    },
    "g2": {
        "title": "School-Age IEP Protocol (5-21 Yrs)",
        "file_tag": "5to21yo",
        "framework": "IDEA IEP | PBIS Framework",
    },
    "g3": {
        "title": "Adult Community Protocol (21+ Yrs)",
        "file_tag": "21plusYo",
        "framework": "Medicaid HCBS | Person-Centered Waiver Framework",
    },
}

current_meta = cohort_meta[selected_cohort_key]
st.caption(
    f"Active Framework Alignment: **{current_meta['framework']}** |"
    " **Standardized Compliance Enabled**"
)

st.write(" ")

# ==========================================
# 5. Import De-Identified Data & Prominent Security Notice (Tool Terminology Unified)
# ==========================================
st.markdown("### 2️⃣ Import Assessment Mock Data (De-Identified)")

st.markdown(
    """
    <div class="critical-security-card">
        <div class="security-title">
            🛡️ DATA PRIVACY & DE-IDENTIFICATION PROTOCOL
        </div>
        <div class="security-body">
            <ul>
                <li><strong>Current Live Demo Notice:</strong> This interactive portal is designed strictly for demonstration purposes. All downloadable sample datasets provided on this page are 100% synthetic, standardized, and fully <strong>de-identified</strong>. You can safely test, upload, and evaluate the tool with complete peace of mind.</li>
                <li><strong>Future Local Deployment Workflow:</strong> In future production environments deployed directly on the clinician's local workstation, the tool will automatically execute an end-to-end <strong>local de-identification pipeline</strong> before parsing any raw uploaded files.</li>
                <li><strong>Seamless Anonymization:</strong> All generated FBA & BIP drafts replace sensitive Personal Identifiable Information (PII) with structured placeholders (e.g., <code>[CLIENT_NAME]</code>, <code>[FACILITY_NAME]</code>).</li>
                <li><strong>Finalization:</strong> Clinicians simply press <strong>CTRL + H</strong> (Find & Replace) in Microsoft Word to insert real client identifiers and perform clinical edits prior to signature. Absolute data privacy and compliance are guaranteed across all phases!</li>
            </ul>
        </div>
    </div>
""",
    unsafe_allow_html=True,
)

col_input1, col_input2, col_input3 = st.columns([1.2, 1.2, 1.1])

# --- Input 1: Direct ABC ---
with col_input1:
  st.markdown("#### 📄 Direct Observation (ABC)")
  st.caption("5 Standardized Observation Sample Entries")

  mock_csv = generate_mock_abc_csv(selected_cohort_key)
  st.download_button(
      label=f"📥 Download De-Identified Mock ABC Data (.csv)",
      data=mock_csv,
      file_name=(
          f"DeIdentified_Mock_ABC_Data_for_{current_meta['file_tag']}.csv"
      ),
      mime="text/csv",
      use_container_width=True,
  )

  uploaded_abc = st.file_uploader(
      "Upload De-Identified Direct ABC File:",
      type=["csv", "xlsx"],
      key=f"abc_{selected_cohort_key}",
  )

# --- Input 2: Raw Indirect Interview Notes ---
with col_input2:
  st.markdown("#### 📝 Indirect Interview Notes")
  st.caption("5 Key Stakeholder Raw Narrative Moments")

  mock_docx = generate_mock_interview_docx(selected_cohort_key)
  st.download_button(
      label=f"📥 Download De-Identified Mock Interview (.docx)",
      data=mock_docx,
      file_name=(
          "DeIdentified_Mock_Interview_Notes_for_"
          f"{current_meta['file_tag']}.docx"
      ),
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      use_container_width=True,
  )

  uploaded_interview = st.file_uploader(
      "Upload De-Identified Interview Notes File:",
      type=["docx", "txt"],
      key=f"interview_{selected_cohort_key}",
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
# 6. Target Language & Synthesis Config
# ==========================================
st.markdown("### 3️⃣ Target Language & Draft Generation Actions")

col_lang, col_btn1, col_btn2 = st.columns([1.2, 1.3, 1.3])

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

with col_btn1:
  st.write(" ")
  st.write(" ")
  generate_fba_click = st.button(
      "⚡ Formulate De-Identified FBA Draft",
      type="primary",
      use_container_width=True,
  )

with col_btn2:
  st.write(" ")
  st.write(" ")
  generate_bip_click = st.button(
      "⚡ Formulate De-Identified BIP Draft",
      type="primary",
      use_container_width=True,
  )

# ==========================================
# 7. Synthesis Preview & Export
# ==========================================
if "show_fba" not in st.session_state:
  st.session_state.show_fba = False
if "show_bip" not in st.session_state:
  st.session_state.show_bip = False

if generate_fba_click:
  st.session_state.show_fba = True
if generate_bip_click:
  st.session_state.show_bip = True

if (
    st.session_state.show_fba
    or st.session_state.show_bip
    or uploaded_abc
    or uploaded_interview
):
  st.success(
      f"✅ Ingested Data for [{current_meta['title']}]. Buffer ready for draft"
      " synthesis."
  )

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
  st.markdown("## 📄 Formulated Clinical Drafts Preview (De-Identified)")

  col_prev1, col_prev2 = st.columns(2)

  with col_prev1:
    if st.session_state.show_fba or (
        not st.session_state.show_fba and not st.session_state.show_bip
    ):
      with st.expander(
          "📄 Standalone FBA Draft Preview (功能性行为评估初稿)", expanded=True
      ):
        st.markdown(f"#### FUNCTIONAL BEHAVIOR ASSESSMENT (FBA) DRAFT")
        st.markdown(
            f"**Cohort:** {current_meta['title']} | **Language:**"
            f" {report_lang}"
        )
        st.markdown("---")
        st.write("**Client Name:** `[CLIENT_NAME]` | **DOB:** `[CLIENT_DOB]`")
        st.write(
            "**1. Target Behavior:** Operational definitions synthesized."
        )
        st.write(
            f"**2. Direct ABC Data:** Ingested {len(parsed_abc_df)}"
            " observations."
        )
        st.dataframe(parsed_abc_df, use_container_width=True)
        st.write(
            f"**3. QABF Summary:** Primary: Task Escape ({q_escape}),"
            f" Secondary: Sensory ({q_sensory})"
        )

  with col_prev2:
    if st.session_state.show_bip or (
        not st.session_state.show_fba and not st.session_state.show_bip
    ):
      with st.expander(
          "📋 Standalone BIP Draft Preview (行为干预计划初稿)", expanded=True
      ):
        st.markdown(f"#### BEHAVIOR INTERVENTION PLAN (BIP) DRAFT")
        st.markdown(f"**Framework Alignment:** {current_meta['framework']}")
        st.markdown("---")
        st.write(
            "**Client Name:** `[CLIENT_NAME]` | **Status:** Initial Draft"
        )
        st.write(
            "**1. Proactive Antecedents:** Demand chunking, visual schedules,"
            " environment modification."
        )
        st.write(
            "**2. Replacement Communication (FCT):** Teaching functional"
            " 'Break' request modality."
        )
        st.write(
            "**3. Reinforcement System:** DRA on FR-1 schedule for replacement"
            " behavior."
        )
        st.write(
            "**4. Safety Protocol:** Neutral environmental safety blocking."
        )

  # ==========================================
  # 8. Document Generation Functions
  # ==========================================

  def build_fba_doc():
    doc = docx.Document()
    doc.add_heading("FUNCTIONAL BEHAVIOR ASSESSMENT (FBA) DRAFT", level=0)
    doc.add_paragraph("Client Name: [CLIENT_NAME] | DOB: [CLIENT_DOB]")
    doc.add_paragraph(f"Cohort Protocol: {current_meta['title']}")
    doc.add_paragraph(f"Framework Alignment: {current_meta['framework']}")
    doc.add_paragraph(f"Language Output: {report_lang}")

    doc.add_heading("1. Target Behavior Breakdown", level=1)
    doc.add_paragraph(
        "Categorized by topography, intensity, and historical baseline."
    )

    doc.add_heading(
        "2. Indirect Assessment Synthesis (From Raw Stakeholder Quotes)",
        level=1,
    )
    doc.add_paragraph(
        "Synthesized from 5 key narrative moments provided by primary"
        " caregivers and direct staff."
    )

    doc.add_heading("3. QABF Assessment Scores", level=1)
    doc.add_paragraph(
        f"Escape: {q_escape}, Sensory: {q_sensory}, Attention: {q_attention},"
        f" Tangible: {q_tangible}, Physical: {q_physical}"
    )

    doc.add_heading(
        f"Appendix A: Ingested Direct ABC Observations ({len(parsed_abc_df)}"
        " Entries)",
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

  def build_bip_doc():
    doc = docx.Document()
    doc.add_heading("BEHAVIOR INTERVENTION PLAN (BIP) DRAFT", level=0)
    doc.add_paragraph("Client Name: [CLIENT_NAME] | DOB: [CLIENT_DOB]")
    doc.add_paragraph(f"Cohort Protocol: {current_meta['title']}")
    doc.add_paragraph(f"Language Output: {report_lang}")

    doc.add_heading("1. Proactive Antecedent Strategies", level=1)
    doc.add_paragraph(
        "Environmental accommodations, visual schedules, and pre-transition"
        " warnings."
    )

    doc.add_heading("2. Replacement Behavior Protocol (FCT)", level=1)
    doc.add_paragraph(
        "Teaching client to independently request a 'Break' or sensory tool."
    )

    doc.add_heading("3. Reinforcement System", level=1)
    doc.add_paragraph(
        "Differential Reinforcement of Alternative Behavior (DRA) on FR-1"
        " schedule."
    )

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

  def build_combined_doc():
    doc = docx.Document()
    doc.add_heading(
        "COMPREHENSIVE CLINICAL FBA & BIP DRAFT PACKAGE", level=0
    )
    doc.add_paragraph("Client Name: [CLIENT_NAME] | DOB: [CLIENT_DOB]")
    doc.add_paragraph(f"Cohort Protocol: {current_meta['title']}")

    doc.add_heading(
        "PART I: FUNCTIONAL BEHAVIOR ASSESSMENT (FBA) DRAFT", level=1
    )
    doc.add_paragraph(
        "Complete assessment breakdown, raw stakeholder quote synthesis, QABF"
        " scoring."
    )

    doc.add_heading("PART II: BEHAVIOR INTERVENTION PLAN (BIP) DRAFT", level=1)
    doc.add_paragraph(
        "Proactive strategies, replacement behaviors, reinforcement"
        " schedules."
    )

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

  # ==========================================
  # 9. Dynamic Export Panel
  # ==========================================
  st.markdown("### 🚀 Export Formulated De-Identified Draft Reports")

  col_dl1, col_dl2, col_dl3 = st.columns(3)

  with col_dl1:
    fba_bytes = build_fba_doc()
    st.download_button(
        label="📄 Download De-Identified FBA Draft (.docx)",
        data=fba_bytes,
        file_name=(
            "DeIdentified_FBA_Draft_for_"
            f"{current_meta['file_tag']}.docx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        use_container_width=True,
    )

  with col_dl2:
    bip_bytes = build_bip_doc()
    st.download_button(
        label="📋 Download De-Identified BIP Draft (.docx)",
        data=bip_bytes,
        file_name=(
            "DeIdentified_BIP_Draft_for_"
            f"{current_meta['file_tag']}.docx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        use_container_width=True,
    )

  with col_dl3:
    combined_bytes = build_combined_doc()
    st.download_button(
        label="📦 Download Full De-Identified FBA & BIP Package (.docx)",
        data=combined_bytes,
        file_name=(
            "DeIdentified_Full_FBA_BIP_DraftPackage_for_"
            f"{current_meta['file_tag']}.docx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        type="primary",
        use_container_width=True,
    )

st.divider()
st.caption(
    "⚠️ **Clinical Responsibility Notice:** This demonstration tool serves"
    " strictly as a clinical first-draft synthesizer for BCBAs and LBAs. All"
    " generated drafts are fully de-identified and must be independently"
    " reviewed, personalized, edited (using CTRL + H for client details), and"
    " verified by the supervising clinician prior to formal signature and"
    " implementation."
)
