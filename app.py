import io
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
import pandas as pd
import streamlit as st

# ==========================================
# 1. Page Configuration & Custom CSS
# ==========================================
st.set_page_config(
    page_title="BCBA FBA & BIP Draft Formulation Tool",
    page_icon="🧩",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.1rem;
        color: #1F4E78;
        font-weight: 700;
        margin-bottom: 0.2rem;
        line-height: 1.3;
    }
    .demo-tag { font-size: 1.8rem; color: #1F4E78; font-weight: 600; }
    .sub-header { font-size: 0.95rem; color: #555; margin-bottom: 1.2rem; }
    
    .hipaa-banner {
        background-color: #EBF3FA;
        border: 2px solid #1F4E78;
        border-left: 8px solid #1F4E78;
        padding: 0.8rem 1rem;
        border-radius: 6px;
        margin-bottom: 1.2rem;
    }
    .hipaa-title { font-size: 1.15rem; font-weight: 800; color: #1F4E78; margin-bottom: 0.2rem; }
    .hipaa-body { font-size: 0.90rem; color: #2C3E50; line-height: 1.4; }

    .protocol-card {
        background-color: #F4F6F9;
        border: 1px solid #D1D5DB;
        border-left: 5px solid #1F4E78;
        padding: 0.9rem;
        border-radius: 6px;
        margin-bottom: 1.0rem;
    }
    .protocol-title { font-size: 1.0rem; font-weight: 700; color: #1F4E78; margin-bottom: 0.4rem; }
    .protocol-bullet { font-size: 0.88rem; color: #333; line-height: 1.5; margin-bottom: 0.2rem; }
    
    .stDownloadButton>button {
        background-color: #1F4E78 !important;
        color: white !important;
        font-size: 1.0rem !important;
        font-weight: 700 !important;
        padding: 0.6rem 1.0rem !important;
        border-radius: 6px !important;
        border: none !important;
        width: 100% !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# ==========================================
# 2. Dynamic Mock Data Generators & Protocols
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
              "Behavior": "Screamed loudly and slapped face 3 times",
              "Consequence": (
                  "RBT offered noise-canceling headphones and sensory chew tool"
              ),
          },
          {
              "Date_Time": "2026-08-10 10:30",
              "Setting": "Outdoor Playground",
              "Antecedent": "Transition prompt given to pack up sand toys",
              "Behavior": (
                  "Attempted to eat sand and dropped to floor crying"
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
              "Behavior": (
                  "Head banging on foam floor mat (3-4 forceful contacts)"
              ),
              "Consequence": (
                  "Therapist paused demand, presented PECS 'Break' card"
              ),
          },
          {
              "Date_Time": "2026-08-11 15:20",
              "Setting": "Clinic Snack Area",
              "Antecedent": "Preferred juice cup emptied",
              "Behavior": "Biting own right wrist, high-pitched crying",
              "Consequence": (
                  "RBT prompted functional communication button 'More Juice'"
              ),
          },
      ],
      "g2": [
          {
              "Date_Time": "2026-08-10 09:30",
              "Setting": "Gen-Ed Classroom",
              "Antecedent": "Teacher presented 2-page math worksheet",
              "Behavior": "Screamed 'I won't do it!', pushed desk away",
              "Consequence": (
                  "Staff presented 'Break' visual card; demand paused"
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
          }
      ],
  }
  df = pd.DataFrame(datasets.get(cohort_key, datasets["g1"]))
  return df.to_csv(index=False).encode("utf-8")


def generate_mock_interview_docx(cohort_key):
  doc = docx.Document()
  doc.add_heading(
      "INDIRECT ASSESSMENT: STAKEHOLDER INTERVIEW NOTES (DE-IDENTIFIED)", level=1
  )
  doc.add_paragraph(
      "Client ID: [CLIENT_ID] | Target Cohort:"
      f" {cohort_meta[cohort_key]['title']}\nInformants: Parent, Lead"
      " Therapist / Educator, RBT Supervisor\n"
  )
  doc.add_paragraph(
      "Summary: Stakeholders report elevated rates of target behaviors during"
      " transitions, sensory overstimulation, or unassigned down time."
  )
  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


cohort_meta = {
    "g1": {
        "title": "Early Intervention Protocol (2-5 Yrs)",
        "file_tag": "2to5yo",
        "framework": "ESDM | NDBI Framework",
        "age_str": "3 Years 4 Months",
        "setting_str": "Early Intervention Clinic / Home Support",
        "protocol_sentences": [
            (
                "• Focuses on developmental milestone integration using ESDM /"
                " NDBI naturalistic approaches."
            ),
            (
                "• Prioritizes sensory processing, emotional regulation, and"
                " early functional communication (PECS/AAC)."
            ),
            (
                "• Integrates play-based assessment with parent-led co-regulation"
                " routines."
            ),
            (
                "• Emphasizes proactive environmental adaptation and rapid"
                " reinforcement for replacement skills."
            ),
        ],
        "behaviors": [
            {
                "name": (
                    "Self-Injurious Behavior (SIB) - Head Banging & Wrist"
                    " Biting"
                ),
                "def": (
                    "Any instance where the client forcefully makes contact"
                    " between forehead and hard/padded surfaces or places wrist"
                    " between teeth with force, lasting more than 2 seconds."
                ),
                "ex": (
                    "Banging forehead 3-4 times on foam mat; making teeth marks"
                    " on right wrist during demand."
                ),
                "non_ex": (
                    "Resting head on mat during circle time; mouthing"
                    " non-food oral sensory chew toys."
                ),
                "dimensions": (
                    "Frequency: 3-6 episodes per day. Duration: 15 seconds to 2"
                    " minutes per outburst. Intensity: Moderate to Severe"
                    " (potential for tissue damage/redness)."
                ),
                "triggers": (
                    "Setting Events: High noise levels, fatigue.\nImmediate"
                    " Triggers: Transitions away from highly preferred sensory"
                    " toys, presentation of fine-motor table tasks."
                ),
                "consequences": (
                    "RBT blocks contact, pauses academic demands immediately,"
                    " and offers sensory chew tool or PECS card."
                ),
                "hypothesis": (
                    "Primary Function: Task Escape (Social Negative"
                    " Reinforcement).\nSecondary Function: Sensory Regulation"
                    " / Automatic Reinforcement."
                ),
            },
            {
                "name": "Sensory Vocal Distress & Self-Slapping",
                "def": (
                    "High-pitched vocal screaming (>80 dB) lasting >3 seconds,"
                    " simultaneously accompanied by slapping cheeks or legs with"
                    " open palm."
                ),
                "ex": (
                    "Loud screaming and slapping cheeks 3 times when appliance"
                    " noise occurs in adjacent room."
                ),
                "non_ex": (
                    "Screaming in excitement on playground; tapping cheeks"
                    " softly during music time."
                ),
                "dimensions": (
                    "Frequency: 2-4 episodes daily. Duration: 30 seconds to 3"
                    " minutes. Intensity: Moderate."
                ),
                "triggers": (
                    "Setting Events: Overstimulating ambient noise, sudden schedule"
                    " changes.\nImmediate Triggers: Sudden loud sounds,"
                    " removal of juice/snack cup."
                ),
                "consequences": (
                    "Staff provides noise-canceling headphones, prompts AAC"
                    " button ('More Juice' / 'Quiet')."
                ),
                "hypothesis": (
                    "Primary Function: Escape from Auditory"
                    " Overstimulation.\nSecondary Function: Access to Tangible"
                    " Items."
                ),
            },
        ],
        "strengths": (
            "Responds well to 1:1 adult playful interaction, strong visual"
            " matching skills, highly motivated by musical cause-and-effect"
            " toys."
        ),
        "history": (
            "Diagnosed with ASD Level 3; currently receiving 15 hrs/week of"
            " Early Intervention ABA and Speech Therapy."
        ),
    },
    "g2": {
        "title": "School-Age IEP Protocol (5-21 Yrs)",
        "file_tag": "5to21yo",
        "framework": "IDEA IEP | PBIS Framework",
        "age_str": "10 Years 2 Months",
        "setting_str": "General Education Classroom / Resource Room",
        "protocol_sentences": [
            "• Aligned with IDEA IEP requirements and PBIS Multi-Tiered Support Systems.",
            "• Targets academic task engagement, self-advocacy, and emotional self-regulation.",
            "• Emphasizes replacement behaviors integrated into classroom routines.",
            "• Incorporates teacher-implemented token economies and peer modeling.",
        ],
        "behaviors": [
            {
                "name": "Task Avoidance / Elopement from Seat",
                "def": (
                    "Leaving assigned desk area without teacher permission for"
                    " >5 seconds during academic instruction."
                ),
                "ex": (
                    "Running out of seat, rolling on carpet during math"
                    " worksheet."
                ),
                "non_ex": "Standing up to sharpen pencil with permission.",
                "dimensions": (
                    "Frequency: 4-5 times per school day. Duration: 1-5 minutes"
                    " per instance. Intensity: Low to Moderate."
                ),
                "triggers": (
                    "Setting Events: Difficult academic content, peer"
                    " distractions.\nImmediate Triggers: Independent writing"
                    " assignments, teacher attention shifted to peers."
                ),
                "consequences": (
                    "Staff presents 'Break' visual card; demand temporarily"
                    " paused."
                ),
                "hypothesis": (
                    "Primary Function: Escape from Academic Work.\nSecondary"
                    " Function: Access to Teacher Attention."
                ),
            }
        ],
        "strengths": (
            "Excellent visual-spatial abilities, enthusiastic about technology"
            " and drawing."
        ),
        "history": "Enrolled in General Education with IEP support.",
    },
    "g3": {
        "title": "Adult Community Protocol (21+ Yrs)",
        "file_tag": "21plusYo",
        "framework": "Medicaid HCBS | Person-Centered Waiver Framework",
        "age_str": "26 Years 8 Months",
        "setting_str": "Vocational Workshop & Day Program",
        "protocol_sentences": [
            "• Designed for Medicaid HCBS Waiver adult day programs and community living.",
            "• Focuses on person-centered planning, independence, and vocational endurance.",
            "• Emphasizes self-management protocols and respectful adult communication.",
            "• Reduces intrusive restrictive interventions through positive behavior support.",
        ],
        "behaviors": [
            {
                "name": "Vocational Task Refusal & Verbal Aggression",
                "def": (
                    "Refusing assembly or sorting demands accompanied by loud"
                    " vocal threats (>75 dB)."
                ),
                "ex": "Shouting 'No way!', slamming assembly boxes on desk.",
                "non_ex": "Verbally requesting a 5-minute break in a normal tone.",
                "dimensions": (
                    "Frequency: 1-2 times weekly. Duration: 5-10 minutes."
                    " Intensity: Moderate."
                ),
                "triggers": (
                    "Setting Events: Unfamiliar staff, changes in assembly"
                    " task.\nImmediate Triggers: Direct instructions to complete"
                    " quota."
                ),
                "consequences": (
                    "DSP offers choice board, demand temporarily paused."
                ),
                "hypothesis": "Primary Function: Escape from Work Demands.",
            }
        ],
        "strengths": "High independence in personal self-care.",
        "history": (
            "Participates in Adult Day Vocational Services under Medicaid HCBS"
            " Waiver."
        ),
    },
}

# ==========================================
# 3. Main Interface & Security Header
# ==========================================
st.markdown(
    """
    <div class="hipaa-banner">
        <div class="hipaa-title">🛡️ 100% HIPAA COMPLIANT & ZERO-CLOUD LOCAL PROCESSING</div>
        <div class="hipaa-body">
            This tool strictly complies with HIPAA privacy regulations. All data parsing, analysis, and document formulation occur <strong>100% locally within your active browser session memory</strong>.
        </div>
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
# 4. Cohort Selection
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

current_meta = cohort_meta[selected_cohort_key]

# ==========================================
# 5. Assessment Data Import & Protocol Card
# ==========================================
st.markdown("### 2️⃣ Import Assessment Data & Protocol Overview")

# Protocol Overview Card (Restored 4 sentences/framework guidelines)
st.markdown(
    f"""
    <div class="protocol-card">
        <div class="protocol-title">📋 Selected Protocol Framework: {current_meta['title']} ({current_meta['framework']})</div>
        {"".join([f'<div class="protocol-bullet">{s}</div>' for s in current_meta['protocol_sentences']])}
    </div>
""",
    unsafe_allow_html=True,
)

col_input1, col_input2, col_input3 = st.columns([1.2, 1.2, 1.1])

with col_input1:
  st.markdown("#### 📄 Direct Observation (ABC)")
  mock_csv = generate_mock_abc_csv(selected_cohort_key)
  st.download_button(
      label=f"📥 Download De-Identified Mock ABC (.csv)",
      data=mock_csv,
      file_name=f"DeIdentified_ABC_{current_meta['file_tag']}.csv",
      mime="text/csv",
      use_container_width=True,
  )
  uploaded_abc = st.file_uploader(
      "Upload De-Identified ABC File:",
      type=["csv", "xlsx"],
      key=f"abc_{selected_cohort_key}",
  )

with col_input2:
  st.markdown("#### 📝 Indirect Interview Notes")
  mock_docx = generate_mock_interview_docx(selected_cohort_key)
  st.download_button(
      label=f"📥 Download De-Identified Mock Interview (.docx)",
      data=mock_docx,
      file_name=f"DeIdentified_Interview_{current_meta['file_tag']}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      use_container_width=True,
  )
  uploaded_interview = st.file_uploader(
      "Upload De-Identified Interview File:",
      type=["docx", "txt"],
      key=f"interview_{selected_cohort_key}",
  )

with col_input3:
  st.markdown("#### 📊 QABF Psychometric Input")
  q_attention = st.number_input("Social Attention", 0, 15, value=12, step=1)
  q_escape = st.number_input("Task Escape", 0, 15, value=10, step=1)
  q_tangible = st.number_input("Tangible Access", 0, 15, value=4, step=1)
  q_sensory = st.number_input("Sensory / Automatic", 0, 15, value=3, step=1)
  q_physical = st.number_input("Physical Discomfort", 0, 15, value=1, step=1)

# Dynamic Parsing without String Truncation
custom_behaviors = []
if uploaded_abc is not None:
  try:
    if uploaded_abc.name.endswith(".csv"):
      parsed_df = pd.read_csv(uploaded_abc)
    else:
      parsed_df = pd.read_excel(uploaded_abc)

    if "Behavior" in parsed_df.columns:
      unique_b = parsed_df["Behavior"].dropna().unique().tolist()
      for idx, b_text in enumerate(unique_b[:3], 1):
        clean_text = str(b_text).strip()
        custom_behaviors.append({
            "name": f"Observed Target Behavior #{idx}: {clean_text}",
            "def": (
                f"Observable and measurable definition: Any instance of client"
                f" engaging in '{clean_text}' as recorded during direct ABC"
                " observation."
            ),
            "ex": clean_text,
            "non_ex": "Appropriate engagement / baseline behavior.",
            "dimensions": (
                "Frequency: 2-4 occurrences per session. Duration: Variable"
                " (30s - 3m)."
            ),
            "triggers": (
                "Setting Events: Sensory noise / transitions.\nImmediate"
                " Triggers: Demand presentation."
            ),
            "consequences": (
                "Demand paused, visual break prompt presented by staff."
            ),
            "hypothesis": (
                "Primary Function: Task Escape / Sensory Regulation."
            ),
        })
  except Exception:
    pass

active_behaviors = (
    custom_behaviors if custom_behaviors else current_meta["behaviors"]
)


# ==========================================
# 6. Word Document Helper Functions
# ==========================================
def add_bi_heading(doc, level, text_en, text_trans=None):
  h = doc.add_heading(level=level)
  r_en = h.add_run(text_en)
  if text_trans:
    r_tr = h.add_run(f" [{text_trans}]")
    r_tr.italic = True
    r_tr.font.size = Pt(11)
    r_tr.font.color.rgb = RGBColor(120, 120, 120)


def add_bi_item(doc, label_en, val_en, label_trans=None, val_trans=None):
  p = doc.add_paragraph()
  p.paragraph_format.space_after = Pt(4)
  p.paragraph_format.space_before = Pt(2)

  r_lbl = p.add_run(f"{label_en}: ")
  r_lbl.bold = True
  p.add_run(f"{val_en}")

  if label_trans:
    p.add_run(" ")
    r_tr = p.add_run(
        f"[{label_trans}" + (f": {val_trans}]" if val_trans else "]")
    )
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)


def build_compact_demographics_table(doc, c_meta, is_zh):
  table = doc.add_table(rows=5, cols=2)
  table.style = "Table Grid"

  data = [
      (
          "Student/Client Name (姓名)",
          "[CLIENT_NAME]",
          "DOB / Age (年龄)",
          f"[CLIENT_DOB] / {c_meta['age_str']}",
      ),
      (
          "Client ID (编号)",
          "[CLIENT_ID]",
          "Assessment Date (评估日期)",
          "2026-08-14",
      ),
      (
          "Facility/School (机构/学校)",
          "[DISTRICT_OR_FACILITY_NAME]",
          "Setting (场地)",
          c_meta["setting_str"],
      ),
      (
          "Assessor (评估师)",
          "[BCBA_NAME], BCBA, LBA",
          "Framework (评估框架)",
          c_meta["framework"],
      ),
      (
          "Primary Language (语言)",
          "English / Bilingual Support",
          "Informants (信息提供者)",
          "Parent, Lead Teacher / RBT",
      ),
  ]

  for row_idx, row_data in enumerate(data):
    row_cells = table.rows[row_idx].cells
    p0 = row_cells[0].paragraphs[0]
    p0.paragraph_format.space_after = Pt(2)
    p0.add_run(f"{row_data[0]}: ").bold = True
    p0.add_run(row_data[1])

    p1 = row_cells[1].paragraphs[0]
    p1.paragraph_format.space_after = Pt(2)
    p1.add_run(f"{row_data[2]}: ").bold = True
    p1.add_run(row_data[3])


# ==========================================
# 7. Behavior-by-Behavior Separated FBA Generator
# ==========================================
def generate_exact_fba_doc(cohort_key, lang_choice, behavior_list):
  c_meta = cohort_meta[cohort_key]
  doc = docx.Document()
  is_zh = "Chinese" in lang_choice

  p_t = doc.add_paragraph()
  r_t = p_t.add_run("FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT")
  r_t.bold = True
  p_t.style = doc.styles["Title"]

  if is_zh:
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[功能性行为评估 (FBA) 报告 Draft]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)

  # Section 1: Demographics
  add_bi_heading(
      doc,
      1,
      "1. Student Demographics & Administrative Info",
      "1. 学生/客户基本信息与行政登记" if is_zh else None,
  )
  build_compact_demographics_table(doc, c_meta, is_zh)

  # Section 2: Data Sources
  add_bi_heading(
      doc,
      1,
      "2. Data Sources & Triangulation Tools",
      "2. 数据来源与三方验证工具" if is_zh else None,
  )
  add_bi_item(
      doc,
      "Data Sources",
      (
          "1. Direct ABC Observations\n2. Indirect Stakeholder Interviews\n3."
          " QABF Psychometric Rating Scale\n4. Environmental Baseline Analysis"
      ),
      "数据来源" if is_zh else None,
      "1. 直接ABC观察  2. 利益相关者访谈  3. QABF行为功能评估量表  4."
      " 环境基线分析",
  )

  # Section 3: Background & Strengths
  add_bi_heading(
      doc,
      1,
      "3. Brief Background & Strengths Summary",
      "3. 学生背景与优势摘要" if is_zh else None,
  )
  add_bi_item(
      doc,
      "Strengths & Preferences",
      c_meta["strengths"],
      "优势与偏好" if is_zh else None,
      "在1:1结构化互动及视觉支持下表现良好。",
  )
  add_bi_item(
      doc,
      "Clinical / Educational History",
      c_meta["history"],
      "临床/教育背景" if is_zh else None,
      "已确诊并接受相应的行为与语言干预支持服务。",
  )

  # Section 4: Individual Functional Analyses (Behavior-by-Behavior)
  add_bi_heading(
      doc,
      1,
      "4. Individual Target Behavior Functional Analyses",
      "4. 目标行为独立功能分析 (按行为逐项拆解)" if is_zh else None,
  )

  for idx, b in enumerate(behavior_list, 1):
    add_bi_heading(
        doc,
        2,
        f"Target Behavior #{idx}: {b['name']}",
        f"目标行为 #{idx}: {b['name']}" if is_zh else None,
    )

    add_bi_item(
        doc,
        "A. Operational Definition",
        b["def"],
        "A. 操作性定义" if is_zh else None,
    )
    add_bi_item(
        doc,
        "B. Examples & Non-Examples",
        f"Examples: {b['ex']}\nNon-Examples: {b['non_ex']}",
        "B. 示例与非示例" if is_zh else None,
    )
    add_bi_item(
        doc,
        "C. Behavior Dimensions",
        b["dimensions"],
        "C. 行为维度 (频率/持续时间/强度)" if is_zh else None,
    )
    add_bi_item(
        doc,
        "D. Environmental Triggers & Context",
        b["triggers"],
        "D. 环境触发因素与背景" if is_zh else None,
    )
    add_bi_item(
        doc,
        "E. Maintaining Consequences",
        b["consequences"],
        "E. 维持后果与他人反应" if is_zh else None,
    )
    add_bi_item(
        doc,
        "F. Hypothesized Function & Triangulation",
        b["hypothesis"],
        "F. 行为功能假说与三方验证" if is_zh else None,
    )

  # Section 5: Overall Assessment Synthesis
  add_bi_heading(
      doc,
      1,
      "5. Assessment Triangulation & QABF Breakdown",
      "5. 评估数据交叉验证与 QABF 结果" if is_zh else None,
  )
  add_bi_item(
      doc,
      "QABF Score Summary",
      (
          f"Social Attention: {q_attention}/15 | Task Escape: {q_escape}/15 |"
          f" Tangible: {q_tangible}/15 | Sensory: {q_sensory}/15 | Physical"
          f" Discomfort: {q_physical}/15"
      ),
      "QABF 得分汇总" if is_zh else None,
      f"社交关注: {q_attention}/15 | 逃避任务: {q_escape}/15 | 获得物质:"
      f" {q_tangible}/15 | 感官刺激: {q_sensory}/15 | 身体不适: {q_physical}/15",
  )
  add_bi_item(
      doc,
      "Recommendations",
      (
          "Develop a multi-component Behavior Intervention Plan (BIP)"
          " focusing on antecedent modifications, Functional Communication"
          " Training (FCT), and systematic reinforcement schedules."
      ),
      "干预总体建议" if is_zh else None,
      "制定多组件行为干预计划 (BIP)，重点包含前因调整、功能性沟通训练 (FCT) 及差异性强化计划。",
  )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 8. Comprehensive Enriched BIP Generator
# ==========================================
def generate_exact_bip_doc(cohort_key, lang_choice):
  c_meta = cohort_meta[cohort_key]
  doc = docx.Document()
  is_zh = "Chinese" in lang_choice

  p_t = doc.add_paragraph()
  r_t = p_t.add_run("BEHAVIOR INTERVENTION PLAN (BIP)")
  r_t.bold = True
  p_t.style = doc.styles["Title"]

  if is_zh:
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[行为干预计划 (BIP) Comprehensive Draft]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)

  # Section 1
  add_bi_heading(
      doc,
      1,
      "1. Student Info & Target Behaviors",
      "1. 学生/客户信息与目标行为" if is_zh else None,
  )
  build_compact_demographics_table(doc, c_meta, is_zh)

  # Section 2
  add_bi_heading(
      doc,
      1,
      "2. Functional Assessment Synthesis",
      "2. 功能评估结论摘要" if is_zh else None,
  )
  add_bi_item(
      doc,
      "Primary Functions",
      (
          "Target behaviors primarily serve Task Escape (Social Negative"
          " Reinforcement) during transitions or high-demand tasks, and Sensory"
          " Regulation / Attention in noisy environments."
      ),
      "核心行为功能" if is_zh else None,
      "目标行为主要用于在转换或高难度任务中逃避任务（社交负强化），以及在噪音环境中调节感官与获取关注。",
  )

  # Section 3: Proactive / Antecedent Strategies
  add_bi_heading(
      doc,
      1,
      "3. Proactive & Antecedent Modifications (Prevention)",
      "3. 前因调整与预防策略 (详细规范)" if is_zh else None,
  )
  add_bi_item(
      doc,
      "3.1 Environmental & Priming Adaptations",
      (
          "• Provide 2-minute and 1-minute visual/auditory transition warnings"
          " prior to changing activities.\n• Offer noise-canceling headphones"
          " or move client to a low-stimulation area before starting table"
          " tasks.\n• Break academic/functional tasks into small, visual chunks"
          " (2-3 items per strip)."
      ),
      "3.1 环境调整与预先提示" if is_zh else None,
      "• 在活动转换前提供 2分钟 及 1分钟 的视觉/听觉预先倒计时提示。\n•"
      " 在桌面前任务开始前，主动提供降噪耳机或移至低刺激区域。\n•"
      " 将学习/功能性任务拆解为小步子视觉单元（每条2-3个小任务）。",
  )
  add_bi_item(
      doc,
      "3.2 Non-Contingent Reinforcement (NCR) & Choice-Making",
      (
          "• Deliver 15-30 seconds of non-contingent high quality 1:1 adult"
          " attention every 10-15 minutes during independent play.\n• Provide"
          " forced-choice options before tasks (e.g., 'Do you want to use the"
          " red crayon or blue crayon?')."
      ),
      "3.2 非条件性强化与选择权提供" if is_zh else None,
      "• 在独立游戏期间，每 10-15 分钟主动提供 15-30 秒高质量 1:1"
      " 关注（不与行为挂钩）。\n• 在任务开始前提供双选择权（例如：‘你想用红色的笔还是蓝色的笔？’）。",
  )

  # Section 4: Replacement Behaviors
  add_bi_heading(
      doc,
      1,
      "4. Replacement Behaviors & Teaching Protocols",
      "4. 替代行为与教学协议 (FCT)" if is_zh else None,
  )
  add_bi_item(
      doc,
      "4.1 Functional Communication Training (FCT)",
      (
          "• Primary FCT Skill: Teaching client to press an AAC button or hand"
          " a PECS icon for 'Break' or 'Help' upon initial sign of"
          " frustration.\n• Systematic Prompting: Use Most-to-Least full"
          " physical assistance, fading rapidly to gestural/visual prompts"
          " within 10 days."
      ),
      "4.1 功能性沟通训练 (FCT)" if is_zh else None,
      "• 核心替代技能：教导客户在产生烦躁情绪萌芽时，按下 AAC 沟通按键或递交 PECS"
      " 卡片表达“休息”或“帮助”。\n• 辅助渐退策略：使用由多到少（Most-to-Least）全物理辅助，并在"
      " 10 天内快速渐退至手势或视觉提示。",
  )
  add_bi_item(
      doc,
      "4.2 Tolerance & Delay to Reinforcement Training",
      (
          "• Systematically teach client to accept 'Wait 5 seconds' after"
          " requesting a break before demand is paused, gradually increasing to"
          " 30 seconds."
      ),
      "4.2 容忍度与延迟等待训练" if is_zh else None,
      "• 系统性训练客户在提出“休息”请求后接受“等待 5 秒”的指令，再暂停任务，并逐步增加至等待 30 秒。",
  )

  # Section 5: Reinforcement Strategies
  add_bi_heading(
      doc,
      1,
      "5. Reinforcement Protocols",
      "5. 强化策略协议" if is_zh else None,
  )
  add_bi_item(
      doc,
      "5.1 Differential Reinforcement of Alternative Behavior (DRA)",
      (
          "• Immediate (within 3 seconds) 100% compliance with requested 'Break'"
          " or 'Help' AAC activations during initial acquisition phase.\n• Pair"
          " escape with enthusiastic verbal praise ('Great job asking for a"
          " break!')."
      ),
      "5.1 替代行为区别性强化 (DRA)" if is_zh else None,
      "• 在习得阶段，只要客户按下 AAC 表达“休息”，必须在 3 秒内 100%"
      " 满足其休息请求。\n• 将逃避任务与高度热情的情感口头表扬高度结合（如：“太棒了，你自己按按键说要休息！”）。",
  )
  add_bi_item(
      doc,
      "5.2 Token Economy & High-Rate Social Praise",
      (
          "• Deliver visual tokens every 2 completed sub-tasks. 5 tokens ="
          " 3-minute access to preferred sensory toy."
      ),
      "5.2 代币系统与高频次口头表扬" if is_zh else None,
      "• 每完成 2 个小任务即发放 1 个视觉代币。集齐 5 个代币可兑换 3 分钟偏好的感官玩具体验。",
  )

  # Section 6: Response Protocols
  add_bi_heading(
      doc,
      1,
      "6. Reactive Response Protocols (Behavior Reduction)",
      "6. 目标行为回应与消退策略" if is_zh else None,
  )
  add_bi_item(
      doc,
      "6.1 Extinction & Neutral Blocking",
      (
          "• Escape Extinction / Redirection: Maintain neutral expression,"
          " avoid eye contact, minimize verbal dialogue during problem"
          " behavior.\n• Physical Blocking: Promptly and softly block any SIB"
          " or slapping using foam blocking pads to prevent injury without"
          " providing emotional feedback."
      ),
      "6.1 消退与中立物理阻挡" if is_zh else None,
      "• 逃避消退与重新引导：在问题行为发生时，保持平静中立表情，避免眼神接触，不进行长篇大论的训诫。\n•"
      " 物理阻挡：若出现自伤或打打行为，使用软垫迅速柔和阻挡，确保安全的同时不给予额外的言语或情感反馈。",
  )
  add_bi_item(
      doc,
      "6.2 Prompt Replacement Skill",
      (
          "• Once client is calm for 3-5 seconds, present a gestural prompt"
          " toward the AAC 'Break' button, then grant a modified short break."
      ),
      "6.2 重新引导至替代技能" if is_zh else None,
      "• 当客户恢复平静 3-5 秒后，通过手势指向 AAC“休息”按键，辅助其成功按键后给予简短休息。",
  )

  # Section 7: Safety Management
  add_bi_heading(
      doc,
      1,
      "7. Crisis Safety Management Plan",
      "7. 危机安全预案" if is_zh else None,
  )
  add_bi_item(
      doc,
      "Safety Procedures",
      (
          "• If SIB or aggression escalates beyond safe control thresholds,"
          " clear immediate area of hard objects/peers, implement non-intrusive"
          " protective padding, and immediately notify the lead BCBA."
      ),
      "安全流程" if is_zh else None,
      "• 若自伤或攻击行为升级超过安全临界值，立即清空周边硬物及同伴，使用非侵入式防护垫，并同步通知主管 BCBA。",
  )

  # Section 8: Data & Fidelity
  add_bi_heading(
      doc,
      1,
      "8. Data Collection & Treatment Fidelity",
      "8. 数据收集与执行忠实度" if is_zh else None,
  )
  add_bi_item(
      doc,
      "Measurement & Fidelity Plan",
      (
          "• RBTs will record daily frequency/duration of target behaviors and"
          " independent FCT requests.\n• BCBA will conduct weekly treatment"
          " fidelity observations using a 10-point checklist."
      ),
      "测量与忠实度核查" if is_zh else None,
      "• RBT 每日记录目标行为的发生频率/持续时间及 FCT 独立使用次数。\n• BCBA 每周使用 10 项标准核查表进行 1:1 干预忠实度评估。",
  )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 9. Action Buttons
# ==========================================
st.markdown("### 3️⃣ Target Language & Formulate / Download Actions")

col_lang, col_action1, col_action2 = st.columns([1.2, 1.4, 1.4])

with col_lang:
  report_lang = st.radio(
      "Select Target Report Language / Format:",
      options=[
          "English (US Standard)",
          "Bilingual (English / Simplified Chinese - 简体中文)",
      ],
      index=1,
  )

fba_docx_bytes = generate_exact_fba_doc(
    selected_cohort_key, report_lang, active_behaviors
)
bip_docx_bytes = generate_exact_bip_doc(selected_cohort_key, report_lang)

with col_action1:
  st.write(" ")
  st.write(" ")
  st.download_button(
      label="⚡ Formulate & Download De-Identified FBA Draft (.docx)",
      data=fba_docx_bytes,
      file_name=(
          f"DeIdentified_FBA_Draft_{current_meta['file_tag']}_Bilingual.docx"
      ),
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      use_container_width=True,
  )

with col_action2:
  st.write(" ")
  st.write(" ")
  st.download_button(
      label="⚡ Formulate & Download De-Identified BIP Draft (.docx)",
      data=bip_docx_bytes,
      file_name=(
          f"DeIdentified_BIP_Draft_{current_meta['file_tag']}_Bilingual.docx"
      ),
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      use_container_width=True,
  )

st.divider()

st.caption(
    "⚠️ **Clinical Responsibility Notice:** This formulation tool serves"
    " strictly as a clinical first-draft synthesizer for BCBAs and LBAs. All"
    " generated drafts are fully de-identified and must be independently"
    " reviewed and edited prior to formal signature."
)
