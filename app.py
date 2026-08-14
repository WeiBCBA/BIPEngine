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

    /* Core Critical Security Notice (4-Line Structured Architecture) */
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
    
    /* High contrast primary action buttons */
    .stDownloadButton>button {
        background-color: #1F4E78 !important;
        color: white !important;
        font-size: 1.05rem !important;
        font-weight: 700 !important;
        padding: 0.8rem 1.2rem !important;
        border-radius: 6px !important;
        border: none !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.12) !important;
        width: 100% !important;
    }
    .stDownloadButton>button:hover {
        background-color: #153552 !important;
        color: #FFFFFF !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# ==========================================
# 2. Dynamic Mock Data Generators (Cohorts)
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


# Dynamic Cohort Meta Definition
cohort_meta = {
    "g1": {
        "title": "Early Intervention Protocol (2-5 Yrs)",
        "file_tag": "2to5yo",
        "framework": "ESDM | NDBI Framework",
        "age_str": "3 Years 4 Months",
        "setting_str": "Early Intervention Clinic / Home Support",
        "behaviors": [
            {
                "name": (
                    "Self-Injurious Behavior (SIB) - Head Banging / Wrist"
                    " Biting"
                ),
                "def": (
                    "Any instance of making forceful contact between head and"
                    " floor/wall or biting own wrist/hand, lasting > 2 seconds."
                ),
                "ex": "Banging forehead on foam mat, biting right wrist.",
                "non_ex": "Gently resting head on pillow, mouthing non-food chew toy.",
            },
            {
                "name": "Sensory Vocal Distress & Slapping",
                "def": (
                    "High-pitched vocal screaming > 85dB accompanied by slapping"
                    " face or body parts when loud environmental sounds occur."
                ),
                "ex": "Screaming and slapping cheeks during blender noise.",
                "non_ex": "Laughter during play, babbling vocalizations.",
            },
        ],
        "strengths": (
            "Responds well to 1:1 adult playful interaction, strong visual"
            " matching skills, enjoys musical cause-and-effect toys."
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
        "behaviors": [
            {
                "name": "Task Avoidance / Elopement from Seat",
                "def": (
                    "Leaving designated desk area for > 5 seconds without"
                    " teacher permission during academic instruction."
                ),
                "ex": "Running out of seat, rolling on carpet during math worksheet.",
                "non_ex": "Standing to sharpen pencil with permission.",
            },
            {
                "name": "Property Destruction (Paper Tearing / Object Throwing)",
                "def": (
                    "Forcefully ripping assigned academic sheets or throwing desk"
                    " items across the room."
                ),
                "ex": "Tearing math packet in half, throwing textbooks.",
                "non_ex": "Accidentally dropping a pencil off the desk.",
            },
        ],
        "strengths": (
            "Excellent visual-spatial abilities, enthusiastic about technology"
            " and drawing, responds well to structured visual schedules."
        ),
        "history": (
            "Enrolled in 4th Grade General Education with IEP support; receives"
            " behavior support and speech services."
        ),
    },
    "g3": {
        "title": "Adult Community Protocol (21+ Yrs)",
        "file_tag": "21plusYo",
        "framework": "Medicaid HCBS | Person-Centered Waiver Framework",
        "age_str": "26 Years 8 Months",
        "setting_str": "Vocational Workshop & Day Program",
        "behaviors": [
            {
                "name": "Vocational Task Refusal & Verbal Aggression",
                "def": (
                    "Refusing assembly/sorting demands accompanied by loud"
                    " vocal threats or slamming workshop items."
                ),
                "ex": (
                    "Shouting 'No way!', slamming assembly boxes onto table."
                ),
                "non_ex": (
                    "Verbally requesting a 5-minute break using self-advocacy"
                    " card."
                ),
            },
            {
                "name": "Disruptive Pacing & Space Blocking",
                "def": (
                    "Pacing rapidly between work aisles or physically blocking"
                    " shared TV/community equipment."
                ),
                "ex": "Standing directly in front of living room TV, blocking path.",
                "non_ex": "Walking around perimeter during scheduled break time.",
            },
        ],
        "strengths": (
            "High independence in personal self-care, skilled in structured"
            " assembly tasks, enjoys community outing trips."
        ),
        "history": (
            "Participates in Adult Day Vocational Services 4 days/week under"
            " Medicaid HCBS Waiver program."
        ),
    },
}


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

current_meta = cohort_meta[selected_cohort_key]
st.caption(
    f"Active Framework Alignment: **{current_meta['framework']}** |"
    " **Standardized Compliance Enabled**"
)

st.write(" ")

# ==========================================
# 5. Import De-Identified Data & Security Notice (Restored 4-Line Architecture)
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
                <li><strong>Live Demo Guidance:</strong> In this online demo phase, please use the pre-loaded standardized <strong>Mock Datasets</strong> below. Do NOT upload real or sensitive client data to the web environment.</li>
                <li><strong>Phase 2 Deployment (Next Phase):</strong> This software will be deployed directly onto the BCBA's local computer, allowing secure processing of local files.</li>
                <li><strong>De-Identification Process:</strong> The engine automatically strips identifiers and replaces them with standardized tags (e.g., <code>[CLIENT_NAME]</code>, <code>[CLIENT_ID]</code>).</li>
                <li><strong>Finalization:</strong> Clinicians simply press <strong>CTRL + H</strong> (Find & Replace) in Microsoft Word to insert actual client identifiers prior to formal signature.</li>
            </ul>
        </div>
    </div>
""",
    unsafe_allow_html=True,
)

col_input1, col_input2, col_input3 = st.columns([1.2, 1.2, 1.1])

with col_input1:
  st.markdown("#### 📄 Direct Observation (ABC)")
  mock_csv = generate_mock_abc_csv(selected_cohort_key)
  st.download_button(
      label=f"📥 Download Mock ABC Data (.csv)",
      data=mock_csv,
      file_name=(
          f"DeIdentified_Mock_ABC_Data_for_{current_meta['file_tag']}.csv"
      ),
      mime="text/csv",
      use_container_width=True,
  )

with col_input2:
  st.markdown("#### 📝 Indirect Interview Notes")
  mock_docx = generate_mock_interview_docx(selected_cohort_key)
  st.download_button(
      label=f"📥 Download Mock Interview (.docx)",
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

with col_input3:
  st.markdown("#### 📊 QABF Psychometric Input")
  q_attention = st.number_input("Social Attention", 0, 15, value=12, step=1)
  q_escape = st.number_input("Task Escape", 0, 15, value=10, step=1)
  q_tangible = st.number_input("Tangible Access", 0, 15, value=4, step=1)
  q_sensory = st.number_input("Sensory / Automatic", 0, 15, value=3, step=1)
  q_physical = st.number_input("Physical Discomfort", 0, 15, value=1, step=1)

st.divider()


# ==========================================
# 6. Bilingual Formatting Helpers (Word Engine)
# ==========================================
def add_bi_heading(doc, level, text_en, text_trans=None):
  h = doc.add_heading(level=level)
  r_en = h.add_run(text_en)
  if text_trans:
    r_tr = h.add_run(f"\n[{text_trans}]")
    r_tr.italic = True
    r_tr.font.size = Pt(Pt(12 if level == 1 else 10))
    r_tr.font.color.rgb = RGBColor(100, 100, 100)


def add_bi_item(
    doc,
    label_en,
    val_en,
    label_trans=None,
    val_trans=None,
    is_heading=False,
):
  p = doc.add_paragraph()
  if is_heading:
    r_head = p.add_run(f"{label_en}")
    r_head.bold = True
    p.style = doc.styles["Heading 2"]
    if val_en:
      p.add_run(f": {val_en}")
  else:
    r_lbl = p.add_run(f"{label_en}: ")
    r_lbl.bold = True
    p.add_run(f"{val_en}")

  if label_trans:
    p_tr = doc.add_paragraph()
    r_tr_lbl = p_tr.add_run(
        f"[{label_trans}" + (f": {val_trans}]" if val_trans else "]")
    )
    r_tr_lbl.italic = True
    r_tr_lbl.font.color.rgb = RGBColor(100, 100, 100)


def build_compact_demographics_table(doc, c_meta, is_zh, is_es):
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
          "Client ID (学号/编号)",
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
          "Informants (信息提供人)",
          "Parent, Lead Teacher / RBT",
      ),
  ]

  for row_idx, row_data in enumerate(data):
    row_cells = table.rows[row_idx].cells
    p0 = row_cells[0].paragraphs[0]
    p0.add_run(f"{row_data[0]}: ").bold = True
    p0.add_run(row_data[1])

    p1 = row_cells[1].paragraphs[0]
    p1.add_run(f"{row_data[2]}: ").bold = True
    p1.add_run(row_data[3])

  doc.add_paragraph()  # spacing


# ==========================================
# 7. Exact 9-Section FBA Document Generator
# ==========================================
def generate_exact_fba_doc(cohort_key, lang_choice):
  c_meta = cohort_meta[cohort_key]
  doc = docx.Document()
  is_zh = "Chinese" in lang_choice
  is_es = "Spanish" in lang_choice

  # Main Document Title
  p_t = doc.add_paragraph()
  r_t = p_t.add_run("FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT")
  r_t.bold = True
  p_t.style = doc.styles["Title"]

  if is_zh:
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[功能性行为评估 (FBA) 报告 Draft]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)
  elif is_es:
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run(
        "[Informe de Evaluación de Conducta Funcional (FBA)]"
    )
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)

  # Section 1: Compact Table Demographics
  add_bi_heading(
      doc,
      1,
      "1. Student Demographics & Administrative Info",
      "1. 学生/客户基本信息与行政登记" if is_zh else "1. Datos Demográficos",
  )
  build_compact_demographics_table(doc, c_meta, is_zh, is_es)

  # Section 2: Data Sources
  add_bi_heading(
      doc,
      1,
      "2. Data Sources & Assessment Tools",
      "2. 数据来源与评估工具" if is_zh else "2. Fuentes de Datos",
  )
  add_bi_item(
      doc,
      "Data Sources (Selected)",
      (
          "Direct ABC Observations, Stakeholder Interviews, Rating Scales"
          " (QABF), Functional Analysis Screening"
      ),
      "数据来源" if is_zh else "Fuentes",
      "直接 ABC 观察, 利益相关者访谈, 评估量表 (QABF), 功能分析筛选",
  )

  # Section 3: Brief Student Background
  add_bi_heading(
      doc,
      1,
      "3. Brief Student Background & Strengths",
      "3. 学生背景与优势" if is_zh else "3. Antecedentes del Estudiante",
  )
  add_bi_item(
      doc,
      "Strengths & Preferences",
      c_meta["strengths"],
      "优势与偏好" if is_zh else "Fortalezas",
      "在1:1互动及结构化视觉支持下表现良好，具有良好的视觉匹配能力。",
  )
  add_bi_item(
      doc,
      "Educational / Clinical History",
      c_meta["history"],
      "教育/临床背景" if is_zh else "Historial Educativo",
      "已确诊并接受相应的行为与语言支持服务。",
  )

  # Section 4: Target Behaviors (Multiple Behaviors Included)
  add_bi_heading(
      doc,
      1,
      "4. Description of Target Behaviors",
      "4. 目标行为描述 (涵盖多项行为)"
      if is_zh
      else "4. Conductas Objetivo",
  )

  for idx, b in enumerate(c_meta["behaviors"], 1):
    add_bi_item(
        doc,
        f"Target Behavior #{idx}",
        b["name"],
        f"目标行为 #{idx}" if is_zh else f"Conducta #{idx}",
    )
    add_bi_item(
        doc,
        "Operational Definition",
        b["def"],
        "操作性定义" if is_zh else "Definición Operacional",
    )
    add_bi_item(
        doc,
        "Examples / Non-examples",
        f"Examples: {b['ex']} | Non-examples: {b['non_ex']}",
        "示例与非示例" if is_zh else "Ejemplos y No Ejemplos",
    )

  # Section 5: Behavior Dimensions
  add_bi_heading(
      doc,
      1,
      "5. Behavior Dimensions (Frequency, Duration, Intensity)",
      "5. 行为维度 (频率、持续时间、强度)"
      if is_zh
      else "5. Dimensiones de la Conducta",
  )
  add_bi_item(
      doc,
      "Frequency & Duration",
      "Occurs an average of 3-5 episodes per session; duration ranges from 30s"
      " to 4 minutes.",
      "发生频率与持续时间" if is_zh else "Frecuencia y Duración",
      "每个观察单元平均发生 3-5 次，持续时间从 30秒 到 4分钟 不等。",
  )
  add_bi_item(
      doc,
      "Intensity & Impact",
      "Moderate to High - Disrupts learning environment and requires immediate"
      " staff redirection.",
      "强度与影响" if is_zh else "Intensidad e Impacto",
      "中度至高度 - 打断学习环境并需要工作人员立即重新引导。",
  )

  # Section 6: Environmental Triggers
  add_bi_heading(
      doc,
      1,
      "6. Environmental Triggers & Context",
      "6. 环境触发因素与背景"
      if is_zh
      else "6. Desencadenantes Ambientales",
  )
  add_bi_item(
      doc,
      "Setting Events (Slow Triggers)",
      "High ambient noise, schedule transitions, fatigue, or unassigned down"
      " time.",
      "背景事件 (慢速触发因素)" if is_zh else "Eventos de Configuración",
      "高环境噪音、日程转换、疲劳或未安排的空闲时间。",
  )
  add_bi_item(
      doc,
      "Antecedents (Immediate Triggers)",
      "Presentation of non-preferred tasks, transition prompts, or withdrawal"
      " of 1:1 adult attention.",
      "前因事件 (快速即时触发因素)" if is_zh else "Antecedentes Inmediatos",
      "呈递非偏好任务、转换提示或撤走 1:1 成人关注。",
  )

  # Section 7: Consequences
  add_bi_heading(
      doc,
      1,
      "7. Maintaining Consequences",
      "7. 维持后果与他人反应" if is_zh else "7. Consecuencias Mantenedoras",
  )
  add_bi_item(
      doc,
      "Staff / Environment Response",
      "Staff approaches client, offers verbal redirection, provides break card,"
      " or pauses demand.",
      "工作人员/环境反应" if is_zh else "Respuesta del Personal",
      "工作人员靠近、提供口头重新引导、呈递暂停卡或暂时暂停任务要求。",
  )

  # Section 8: Function Hypothesis
  add_bi_heading(
      doc,
      1,
      "8. Clinical Hypothesis & Function of Behavior",
      "8. 临床假说与行为功能" if is_zh else "8. Hipótesis Funcional",
  )
  add_bi_item(
      doc,
      "Functional Hypothesis",
      (
          f"When presented with sensory overstimulation or demanding tasks in"
          f" the {c_meta['setting_str']}, the client engages in target"
          f" behaviors to gain attention or escape demands."
      ),
      "临床假说" if is_zh else "Hipótesis Clínica",
      "当遇到刺激或高难度任务时，客户通过目标行为以获取关注或逃避任务。",
  )
  add_bi_item(
      doc,
      "Primary Functions (QABF Breakdown)",
      (
          f"1. Social Attention (Score: {q_attention}/15) | 2. Task Escape"
          f" (Score: {q_escape}/15) | 3. Sensory (Score: {q_sensory}/15)"
      ),
      "主要功能 (QABF得分)" if is_zh else "Funciones Principales",
      f"1. 获取关注 ({q_attention}/15) | 2. 逃避任务 ({q_escape}/15) | 3."
      f" 感官调节 ({q_sensory}/15)",
  )

  # Section 9: Clinical Notes
  add_bi_heading(
      doc,
      1,
      "9. Additional Clinical Recommendations",
      "9. 附加临床建议" if is_zh else "9. Recomendaciones Clínicas",
  )
  add_bi_item(
      doc,
      "Recommendations",
      "Develop immediate BIP focused on proactive antecedent modifications,"
      " functional communication training (FCT), and visual schedules.",
      "干预建议" if is_zh else "Recomendaciones",
      "建议立即制定 BIP，重点关注预防性前因调整、功能性沟通训练 (FCT) 及视觉日程表。",
  )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 8. Exact 9-Section BIP Document Generator
# ==========================================
def generate_exact_bip_doc(cohort_key, lang_choice):
  c_meta = cohort_meta[cohort_key]
  doc = docx.Document()
  is_zh = "Chinese" in lang_choice
  is_es = "Spanish" in lang_choice

  # Main Document Title
  p_t = doc.add_paragraph()
  r_t = p_t.add_run("BEHAVIOR INTERVENTION PLAN (BIP)")
  r_t.bold = True
  p_t.style = doc.styles["Title"]

  if is_zh:
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[行为干预计划 (BIP) Draft]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)
  elif is_es:
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[Plan de Intervención de la Conducta (BIP)]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)

  # Section 1: Compact Table Demographics
  add_bi_heading(
      doc,
      1,
      "1. Student Info & Target Behaviors",
      "1. 学生/客户信息与目标行为" if is_zh else "1. Datos del Estudiante",
  )
  build_compact_demographics_table(doc, c_meta, is_zh, is_es)

  # Section 2: FBA Hypothesis
  add_bi_heading(
      doc,
      1,
      "2. FBA Functional Hypothesis Summary",
      "2. FBA 行为功能假说摘要" if is_zh else "2. Resumen de Hipótesis",
  )
  add_bi_item(
      doc,
      "Functional Summary",
      (
          "Target behaviors function primarily to gain adult attention and"
          " escape overstimulating or difficult tasks."
      ),
      "功能摘要" if is_zh else "Resumen Funcional",
      "目标行为的主要功能为获取关注及逃避高难度或高刺激任务。",
  )

  # Section 3: Antecedent Modifications
  add_bi_heading(
      doc,
      1,
      "3. Antecedent Modifications (Prevention)",
      "3. 前因调整与预防策略"
      if is_zh
      else "3. Modificaciones de Antecedentes",
  )
  add_bi_item(
      doc,
      "Proactive Strategies",
      (
          "• Pre-activity transition warnings and visual schedule count-down.\n•"
          " Provide non-contingent attention check-ins every 10-15 minutes.\n•"
          " Reduce environmental noise / provide noise-canceling headphones"
          " when needed.\n• Chunk tasks into smaller, manageable visual units."
      ),
      "预防性策略" if is_zh else "Estrategias Proactivas",
      "• 转换前的预先提示与视觉倒计时。\n• 每 10-15"
      " 分钟提供非条件性关注与确认。\n• 降低环境噪音，必要时提供降噪耳机。\n•"
      " 将任务拆解为更小的视觉单元。",
  )

  # Section 4: Replacement Behaviors
  add_bi_heading(
      doc,
      1,
      "4. Replacement Behaviors & Teaching Protocols",
      "4. 替代行为与教学协议"
      if is_zh
      else "4. Conductas de Reemplazo",
  )
  add_bi_item(
      doc,
      "Functional Communication Training (FCT)",
      (
          "Client will be taught to press a PECS/AAC 'Break' or 'Help' button"
          " instead of engaging in target behaviors."
      ),
      "功能性沟通训练 (FCT)" if is_zh else "Entrenamiento en Comunicación",
      "教导客户使用 PECS/AAC 沟通按键（表达“休息”或“帮助”）来替代目标行为。",
  )
  add_bi_item(
      doc,
      "Teaching Protocol",
      (
          "• Systematic Prompting & Fading (Errorless Learning).\n• Behavioral"
          " Rehearsal / Role-play during low-stress times."
      ),
      "教学协议" if is_zh else "Protocolo de Enseñanza",
      "• 渐进式提示与消除（无错误学习法）。\n• 在低压力状态下进行行为演练与角色扮演。",
  )

  # Section 5: Reinforcement Strategies
  add_bi_heading(
      doc,
      1,
      "5. Reinforcement Strategies for Replacement Behavior",
      "5. 替代行为的强化策略"
      if is_zh
      else "5. Estrategias de Refuerzo",
  )
  add_bi_item(
      doc,
      "Differential Reinforcement (DRA/DRO)",
      (
          "• Immediate (within 3s) access to requested item/break upon using"
          " replacement card/button.\n• High-rate social praise and visual token"
          " economy system."
      ),
      "区别性强化策略" if is_zh else "Refuerzo Diferencial",
      "• 使用替代沟通工具后，立即（3秒内）满足其休息或物品需求。\n•"
      " 高频次的口头表扬与视觉代币奖励体系。",
  )

  # Section 6: Response Protocols
  add_bi_heading(
      doc,
      1,
      "6. Response Protocols for Target Behavior Reduction",
      "6. 目标行为削减与回应策略" if is_zh else "6. Protocolo de Respuesta",
  )
  add_bi_item(
      doc,
      "Extinction & Redirect Protocols",
      (
          "• Neutral Extinction: Minimize verbal eye contact during behavior"
          " outburst.\n• Physical Block: Promptly block any SIB or property"
          " destruction for safety.\n• Prompt Replacement: Redirect neutrally"
          " to 'Break' card."
      ),
      "消退与重新引导协议" if is_zh else "Extinción y Redirección",
      "• 中立消退：发作期间保持中立，减少眼神与言语接触。\n• 物理阻挡：出于安全迅速阻挡自伤或破坏行为。\n•"
      " 重新引导：中立地重新引导至“休息”卡片。",
  )

  # Section 7: Crisis Safety Plan
  add_bi_heading(
      doc,
      1,
      "7. Crisis Safety Management Plan",
      "7. 危机安全预案" if is_zh else "7. Plan de Crisis y Seguridad",
  )
  add_bi_item(
      doc,
      "Safety Procedures",
      (
          "If target behavior escalates to safety hazard, clear immediate area,"
          " implement padded mats if needed, and contact lead BCBA."
      ),
      "安全流程" if is_zh else "Procedimientos de Seguridad",
      "若行为升级危及安全，立即清空周边区域，必要时使用软垫防护，并通知主管 BCBA。",
  )

  # Section 8: Data Collection
  add_bi_heading(
      doc,
      1,
      "8. Data Collection & Progress Monitoring",
      "8. 数据收集与进度监控" if is_zh else "8. Recolección de Datos",
  )
  add_bi_item(
      doc,
      "Measurement Method",
      (
          "RBT/Staff will record frequency of target behaviors and FCT"
          " independence daily using electronic data sheets."
      ),
      "测量方法" if is_zh else "Método de Medición",
      "工作人员每日通过电子表格记录目标行为发生频率以及 FCT 独立使用的次数。",
  )

  # Section 9: Staff Training
  add_bi_heading(
      doc,
      1,
      "9. Staff Training & Implementation Fidelity",
      "9. 人员培训与执行忠实度" if is_zh else "9. Entrenamiento del Personal",
  )
  add_bi_item(
      doc,
      "Fidelity Training Plan",
      (
          "BCBA provides competency-based modeling and weekly treatment"
          " fidelity checklists."
      ),
      "执行忠实度计划" if is_zh else "Plan de Fidelidad",
      "BCBA 提供基于能力的示范指导，并每周开展干预忠实度核查。",
  )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 9. Action Section: Direct Formulate & Download
# ==========================================
st.markdown("### 3️⃣ Target Language & Formulate / Download Actions")

col_lang, col_action1, col_action2 = st.columns([1.2, 1.4, 1.4])

with col_lang:
  report_lang = st.radio(
      "Select Target Report Language / Format:",
      options=[
          "English (US Standard)",
          "Bilingual (English / Simplified Chinese - 简体中文)",
          "Bilingual (English / Spanish - Español)",
      ],
      index=1,
  )

# Pre-generate cohort-matched docs for direct high-speed download
fba_docx_bytes = generate_exact_fba_doc(selected_cohort_key, report_lang)
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
    " reviewed, personalized, edited (using CTRL + H for client details), and"
    " verified by the supervising clinician prior to formal signature."
)
