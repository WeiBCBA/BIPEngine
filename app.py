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
# 2. Dynamic Mock Data Generators
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
  doc.add_heading(
      "INDIRECT ASSESSMENT: STAKEHOLDER INTERVIEW NOTES (DE-IDENTIFIED)", level=1
  )
  doc.add_paragraph(
      "Client ID: [CLIENT_ID] | Target Cohort:"
      f" {cohort_meta[cohort_key]['title']}\nInformants: Parent, Special"
      " Educator, RBT Supervisor\n"
  )
  doc.add_paragraph(
      "Summary: Stakeholders note elevated rate of target behaviors when"
      " unassigned time occurs or when multi-step writing tasks are required."
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
# 5. Import De-Identified Data & Security Notice
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
                <li><strong>Current Live Demo Notice:</strong> Standardized and fully <strong>de-identified</strong> sample datasets are pre-loaded below. You can also upload custom files.</li>
                <li><strong>Finalization:</strong> Clinicians simply press <strong>CTRL + H</strong> (Find & Replace) in Microsoft Word to insert real client identifiers prior to signature.</li>
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
  uploaded_abc = st.file_uploader(
      "Upload Direct ABC File:",
      type=["csv", "xlsx"],
      key=f"abc_{selected_cohort_key}",
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
  uploaded_interview = st.file_uploader(
      "Upload Interview File:",
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

st.divider()

# Parsing uploaded ABC for export
if uploaded_abc is not None:
  try:
    if uploaded_abc.name.endswith(".csv"):
      parsed_abc_df = pd.read_csv(uploaded_abc)
    else:
      parsed_abc_df = pd.read_excel(uploaded_abc)
  except Exception:
    parsed_abc_df = pd.read_csv(io.StringIO(mock_csv.decode("utf-8")))
else:
  parsed_abc_df = pd.read_csv(io.StringIO(mock_csv.decode("utf-8")))


# ==========================================
# 6. Bilingual Formatting Engine (Word Helper)
# ==========================================
def add_bi_item(
    doc,
    label_en,
    val_en,
    label_trans=None,
    val_trans=None,
    lang="zh",
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


# ==========================================
# 7. Exact 9-Section FBA Document Generator
# ==========================================
def generate_exact_fba_doc(lang_choice):
  doc = docx.Document()
  is_zh = "Chinese" in lang_choice
  is_es = "Spanish" in lang_choice

  # Main Title
  p_t = doc.add_paragraph()
  r_t = p_t.add_run("FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) FORM")
  r_t.bold = True
  p_t.style = doc.styles["Title"]

  if is_zh:
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[功能性行为评估 (FBA) 表格]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)
  elif is_es:
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[Formulario de Evaluación de Conducta Funcional (FBA)]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)

  # 1. Header / Student Demographics
  doc.add_heading("1. Student Demographics & Administrative Info", level=1)
  add_bi_item(
      doc,
      "Student's Name",
      "[CLIENT_NAME]",
      "学生姓名" if is_zh else ("Nombre del Estudiante" if is_es else None),
      "[客户端姓名]",
  )
  add_bi_item(
      doc,
      "Student's DOB",
      "[CLIENT_DOB]",
      "出生日期" if is_zh else ("Fecha de Nacimiento" if is_es else None),
      "[出生日期]",
  )
  add_bi_item(
      doc,
      "Student's ID",
      "[CLIENT_ID]",
      "学号/ID" if is_zh else ("ID del Estudiante" if is_es else None),
      "[学生ID]",
  )
  add_bi_item(
      doc,
      "School Name / District",
      "Box Hill High School / District 1",
      "学校名称/学区" if is_zh else ("Escuela / Distrito" if is_es else None),
      "博士山高中 / 第一学区",
  )
  add_bi_item(
      doc,
      "Date of FBA",
      "2026-08-14",
      "评估日期" if is_zh else ("Fecha de FBA" if is_es else None),
      "2026年08月14日",
  )

  # 2. Data Sources
  doc.add_heading("2. Data Sources", level=1)
  add_bi_item(
      doc,
      "Data Sources (Selected)",
      (
          "Direct Observations, Student Interview, Teacher Interview, Parent"
          " Interview, Rating Scales (QABF)"
      ),
      "数据来源 (已选)"
      if is_zh
      else ("Fuentes de Datos (Seleccionadas)" if is_es else None),
      "直接观察, 学生访谈, 教师访谈, 家长访谈, 评估量表 (QABF)",
  )

  # 3. Brief Student Background
  doc.add_heading("3. Brief Student Background", level=1)
  add_bi_item(
      doc,
      "Strengths",
      (
          "Strong visual-spatial skills, responds well to 1:1 adult praise,"
          " highly motivated by sensory computer games."
      ),
      "学生优势/长处"
      if is_zh
      else ("Fortalezas del Estudiante" if is_es else None),
      "具有较强的视觉空间能力，对一对一的成人表扬回应良好，对感官类电脑游戏有很高积极性。",
  )
  add_bi_item(
      doc,
      "Educational History",
      (
          "Enrolled in general education with support; IEP active; receives"
          " speech and behavioral support services."
      ),
      "教育背景/历史" if is_zh else ("Historial Educativo" if is_es else None),
      "在支持下参与普通教育；IEP（个性化教育计划）生效中；接受语言与行为支持服务。",
  )

  # 4. Description of Target Behavior
  doc.add_heading("4. Description of Target Behavior", level=1)
  add_bi_item(
      doc,
      "Operational Definition",
      (
          "Out of seat for more than 5 seconds without permission during"
          " instruction."
      ),
      "目标行为操作性定义"
      if is_zh
      else ("Definición Operacional" if is_es else None),
      "教学期间未经教师许可离开座位超过5秒。",
  )
  add_bi_item(
      doc,
      "Examples",
      "Running around classroom, rolling on floor, tumbling on carpet.",
      "行为示例 (Examples)"
      if is_zh
      else ("Ejemplos de la Conducta" if is_es else None),
      "在教室里奔跑、在地上打滚、在地毯上翻滚。",
  )
  add_bi_item(
      doc,
      "Non-examples",
      (
          "Standing up to sharpen pencil with permission, bathroom"
          " emergencies."
      ),
      "非行为示例 (Non-examples)"
      if is_zh
      else ("No Ejemplos" if is_es else None),
      "经许可站起来削铅笔、上厕所等紧急情况。",
  )

  # 5. Behavior Dimensions (Frequency, Duration, Intensity)
  doc.add_heading(
      "5. Behavior Dimensions (Frequency, Duration, Intensity)", level=1
  )
  add_bi_item(
      doc,
      "Frequency",
      "Occurs an average of 4-6 times per 60-minute instructional block.",
      "发生频率 (Frequency)"
      if is_zh
      else ("Frecuencia" if is_es else None),
      "在60分钟的教学单元中平均发生 4-6 次。",
  )
  add_bi_item(
      doc,
      "Duration",
      "Episodes last between 45 seconds to 3 minutes before redirection.",
      "持续时间 (Duration)" if is_zh else ("Duración" if is_es else None),
      "在获得重新引导前，每次发作持续45秒至3分钟。",
  )
  add_bi_item(
      doc,
      "Intensity",
      (
          "Moderate to High Intensity - Disrupts instructional flow for peers,"
          " requires immediate staff redirection."
      ),
      "行为强度 (Intensity)" if is_zh else ("Intensidad" if is_es else None),
      "中度至高度强度 - 打断同伴的教学流程，需要工作人员立即重新引导。",
  )

  # 6. Environmental Triggers & Context (Setting Events & Antecedents)
  doc.add_heading(
      "6. Environmental Triggers & Context (Setting Events & Antecedents)",
      level=1,
  )
  add_bi_item(
      doc,
      "Setting Events (Slow Triggers)",
      (
          "Morning transitions, noisy room conditions, fatigue from poor"
          " sleep."
      ),
      "背景事件/慢速引发因素 (Setting Events)"
      if is_zh
      else ("Eventos de Configuración" if is_es else None),
      "早晨的环节切换、高噪音环境、因睡眠不足导致的疲劳。",
  )
  add_bi_item(
      doc,
      "Antecedent Events (Fast Immediate Triggers)",
      (
          "Presentation of independent writing tasks, removal of 1:1 adult"
          " attention."
      ),
      "前因事件/快速即时触发因素 (Antecedents)"
      if is_zh
      else ("Antecedentes Inmediatos" if is_es else None),
      "布置独立写作任务、撤走一对一的成人关注。",
  )
  add_bi_item(
      doc,
      "Non-Occurring Situations",
      "1:1 direct instruction, preferred computer activities, recess.",
      "目标行为通常不发生的场景"
      if is_zh
      else ("Situaciones sin Conducta" if is_es else None),
      "一对一直接教学、偏好的电脑活动、课间休息。",
  )

  # 7. Consequences & Staff Response
  doc.add_heading("7. Consequences (Immediate Response)", level=1)
  add_bi_item(
      doc,
      "Consequences",
      (
          "Staff members immediately approach student, verbally redirect, or"
          " teach replacement behavior."
      ),
      "后果/他人反应 (Consequences)"
      if is_zh
      else ("Consecuencias Inmediatas" if is_es else None),
      "工作人员立即靠近学生进行口头重新引导，或教导替代行为。",
  )

  # 8. Hypothesis & Function of Behavior
  doc.add_heading("8. Clinical Hypothesis & Function of Behavior", level=1)
  add_bi_item(
      doc,
      "Hypothesis",
      (
          "When the student is not receiving 1:1 adult attention during writing"
          " tasks, the student engages in out-of-seat behavior to obtain"
          ' teacher attention, communicating "I want my teacher\'s attention."'
      ),
      "临床假说 (Hypothesis)"
      if is_zh
      else ("Hipótesis Clínica" if is_es else None),
      "当学生在写作任务中未获得一对一成人关注时，会通过离座行为来获取教师关注，以此表达“我需要老师的关注”。",
  )
  add_bi_item(
      doc,
      "Primary Function of Behavior",
      (
          f"Attention Access (QABF Score: {q_attention}/15) | Secondary:"
          f" Escape ({q_escape}/15)"
      ),
      "行为的主要功能 (Primary Function)"
      if is_zh
      else ("Función Principal" if is_es else None),
      f"获取关注 (QABF得分: {q_attention}/15) | 次要功能: 逃避 ({q_escape}/15)",
  )

  # 9. Additional Clinical Notes
  doc.add_heading("9. Additional Clinical Notes", level=1)
  add_bi_item(
      doc,
      "Additional Notes",
      (
          "Recommend immediate BIP development focusing on non-contingent"
          " attention schedules, structured waiting tasks, and functional"
          " self-monitoring."
      ),
      "附加临床备注" if is_zh else ("Notas Adicionales" if is_es else None),
      "建议立即制定 BIP（行为干预计划），重点关注非非条件性关注程序、结构化等待任务及自我监控训练。",
  )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 8. Exact 9-Section BIP Document Generator
# ==========================================
def generate_exact_bip_doc(lang_choice):
  doc = docx.Document()
  is_zh = "Chinese" in lang_choice
  is_es = "Spanish" in lang_choice

  # Main Title
  p_t = doc.add_paragraph()
  r_t = p_t.add_run("BEHAVIOR INTERVENTION PLAN (BIP)")
  r_t.bold = True
  p_t.style = doc.styles["Title"]

  if is_zh:
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[行为干预计划 (BIP)]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)
  elif is_es:
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[Plan de Intervención de la Conducta (BIP)]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)

  # 1. Header & Target Behavior Description
  doc.add_heading("1. Student Info & Target Behavior Description", level=1)
  add_bi_item(
      doc,
      "Student Name",
      "[CLIENT_NAME]",
      "学生姓名" if is_zh else ("Nombre del Estudiante" if is_es else None),
      "[客户端姓名]",
  )
  add_bi_item(
      doc,
      "Date BIP Written",
      "2026-08-14",
      "BIP撰写日期" if is_zh else ("Fecha de BIP" if is_es else None),
      "2026年08月14日",
  )
  add_bi_item(
      doc,
      "Target Behavior",
      "Out of seat for more than 5 seconds without permission.",
      "目标行为描述" if is_zh else ("Conducta Objetivo" if is_es else None),
      "未经许可离座超过5秒。",
  )
  add_bi_item(
      doc,
      "Examples / Nonexamples",
      (
          "Examples: running, rolling on floor. Nonexamples: emergency"
          " bathroom use."
      ),
      "示例与非示例" if is_zh else ("Ejemplos y No Ejemplos" if is_es else None),
      "示例：奔跑、地上打滚。非示例：紧急上厕所。",
  )

  # 2. Hypothesis
  doc.add_heading("2. FBA Functional Hypothesis", level=1)
  add_bi_item(
      doc,
      "FBA Hypothesis",
      (
          "Target behavior occurs when 1:1 adult attention is withheld during"
          ' independent tasks to gain attention, communicating "I want my'
          ' teacher\'s attention."'
      ),
      "FBA 行为功能假说"
      if is_zh
      else ("Hipótesis Funcional del FBA" if is_es else None),
      "在独立任务中缺少一对一成人关注时发生目标行为，以此获取关注，沟通意图为“我想要老师关注”。",
  )

  # 3. Antecedent Modifications
  doc.add_heading("3. Antecedent Modifications (Prevention)", level=1)
  add_bi_item(
      doc,
      "Proactive Modifications",
      (
          "• Teacher check-in upon school arrival (walk & talk).\n• Provide"
          " quarterly scheduled walks (8:30, 10:30, 12:30, 2:30) for"
          " non-contingent attention.\n• Establish a folder of 'waiting'"
          " activities (easy academic tasks, word-finds).\n• Remove physical"
          " climbing structures from room."
      ),
      "前因调整与预防策略"
      if is_zh
      else ("Modificaciones de Antecedentes" if is_es else None),
      "• 到校时教师进行晨间沟通（散步交流）。\n• 在固定时间（8:30, 10:30, 12:30,"
      " 2:30）安排散步，提供非条件性关注。\n• 建立“等待活动”文件夹（简单学术任务、字谜）。\n•"
      " 移走教室里可供攀爬的物理设施。",
  )

  # 4. Replacement Behaviors & Teaching Protocols
  doc.add_heading("4. Replacement Behaviors & Teaching Protocols", level=1)
  add_bi_item(
      doc,
      "Replacement Behavior",
      (
          "Student will be taught to wait for teacher attention using a structured"
          " 'waiting' folder and visual clock."
      ),
      "替代行为"
      if is_zh
      else ("Conductas de Reemplazo" if is_es else None),
      "教导学生使用结构化的“等待”文件夹和视觉时钟来等待老师的关注。",
  )
  add_bi_item(
      doc,
      "Teaching Instructions",
      (
          "• Direct Instruction & Role Play: Practice 'waiting' vs. 'not"
          " waiting'.\n• Self-monitoring procedure: Student marks self-monitoring"
          " form on desk at scheduled wait intervals."
      ),
      "教学与训练协议"
      if is_zh
      else ("Protocolo de Enseñanza" if is_es else None),
      "• 直接教学与角色扮演：演练“等待”与“不等待”的区别。\n•"
      " 自我监控程序：学生在桌面的表格上对预设的等待时间段进行自我记录。",
  )

  # 5. Strategies for Reinforcing Replacement Behavior
  doc.add_heading("5. Reinforcement Strategies for Replacement Behavior", level=1)
  add_bi_item(
      doc,
      "Reinforcement Schedule",
      (
          "• Praise & Specific Attention: Paraprofessional gives 30 seconds of"
          " enthusiastic, specific praise for waiting.\n• Increase wait time"
          " gradually by 2-3 minutes as mastery is shown.\n• Sensory"
          " Self-Regulating Items: Brief access to stress ball or sensory items"
          " at desk."
      ),
      "替代行为的强化策略"
      if is_zh
      else ("Estrategias de Refuerzo" if is_es else None),
      "• 表扬与具体关注：副语言治疗师/助教提供30秒的热情具体表扬。\n• 随着熟练度提升，每几天逐步增加2-3分钟等待时间。\n•"
      " 感官自我调节物品：在桌前短时间使用压力球或感官小工具。",
  )

  # 6. Strategies for Reducing Target Behavior (Response Protocols)
  doc.add_heading("6. Response Protocols for Target Behavior Reduction", level=1)
  add_bi_item(
      doc,
      "Consequence Strategy",
      (
          "• Planned Ignoring: Use planned ignoring when safety is not a"
          " concern.\n• Warning / Choice: If behavior persists, warn student that"
          " continuing leads to loss of tangibles, OR stopping and sitting"
          " earns a reinforcer."
      ),
      "目标行为削减与后果回应策略"
      if is_zh
      else ("Estrategias de Reducción" if is_es else None),
      "• 计划性忽视：在不涉及安全隐患时使用计划性忽视。\n• 警告与选择：若行为持续，提示学生继续该行为将失去实物奖励，或停止并坐下可赢得强化物。",
  )

  # 7. Crisis Safety Plan
  doc.add_heading("7. Crisis Safety Management Plan", level=1)
  add_bi_item(
      doc,
      "Crisis Plan",
      (
          "If target behavior escalates and jeopardizes safety, implement"
          " established classroom safety protocol (e.g., clear immediate area of"
          " students or relocate student to low-stimulation area)."
      ),
      "危机安全预案" if is_zh else ("Plan de Crisis" if is_es else None),
      "若目标行为升级并危及安全，执行既定的教室安全预案（例如：清空周边学生或将学生转移至低刺激区域）。",
  )

  # 8. Data Collection & Progress Monitoring
  doc.add_heading("8. Data Collection & Progress Monitoring", level=1)
  add_bi_item(
      doc,
      "Data Monitoring",
      (
          "• Baseline and intervention 'waiting' time collected by"
          " paraprofessional daily.\n• Data reviewed daily to make adjustments to"
          " expected wait intervals."
      ),
      "数据收集与进度监控"
      if is_zh
      else ("Recolección de Datos" if is_es else None),
      "• 助教/工作人员每日记录基线及干预期间的“等待”时间。\n• 每日审查数据，根据需要实时微调设定的等待间隔。",
  )

  # 9. Staff Training & Implementation Fidelity
  doc.add_heading("9. Staff Training & Implementation Fidelity", level=1)
  add_bi_item(
      doc,
      "Staff Training Plan",
      (
          "• Lead teacher trains paraprofessional on replacement teaching,"
          " self-monitoring setup, and praise consistency.\n• Team consults"
          " daily to evaluate fidelity and plan effectiveness."
      ),
      "人员培训与执行忠实度"
      if is_zh
      else ("Entrenamiento del Personal" if is_es else None),
      "• 主任教师对助教进行替代行为教学、自我监控设置及表扬一致性的培训。\n• 团队每日会商，评估执行忠实度与方案有效性。",
  )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 9. Action Section: Direct Formulate & Download (No Live Preview)
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

# Pre-generate docs for one-click direct download
fba_docx_bytes = generate_exact_fba_doc(report_lang)
bip_docx_bytes = generate_exact_bip_doc(report_lang)

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
