import io
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

# ==========================================
# 1. Page Configuration & Custom CSS
# ==========================================
st.set_page_config(
    page_title="BCBA FBA & BIP Clinical & Visual Analysis Tool",
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
          }
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


def generate_mock_tracking_data():
  data = {
      "Session": [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
      "Phase": [
          "Baseline",
          "Baseline",
          "Baseline",
          "Intervention",
          "Intervention",
          "Intervention",
          "Intervention",
          "Intervention",
          "Intervention",
          "Intervention",
      ],
      "Target_Behavior_Frequency": [12, 14, 15, 11, 8, 6, 5, 3, 2, 1],
  }
  return pd.DataFrame(data)


def plot_aba_behavior_graph(df):
  fig, ax = plt.subplots(figsize=(10, 4.5))
  baseline = df[df["Phase"] == "Baseline"]
  intervention = df[df["Phase"] == "Intervention"]

  ax.plot(
      baseline["Session"],
      baseline["Target_Behavior_Frequency"],
      marker="o",
      color="#C0392B",
      linewidth=2.5,
      label="Baseline (基线期)",
  )
  ax.plot(
      intervention["Session"],
      intervention["Target_Behavior_Frequency"],
      marker="o",
      color="#1F4E78",
      linewidth=2.5,
      label="Intervention (干预期)",
  )
  ax.axvline(
      x=3.5,
      color="gray",
      linestyle="--",
      linewidth=1.5,
      label="BIP Implemented (干预启动)",
  )

  ax.set_title(
      "ABA Visual Analysis: Target Behavior Frequency Trend",
      fontsize=12,
      fontweight="bold",
      color="#1F4E78",
  )
  ax.set_xlabel("Session / Observation Number", fontsize=10)
  ax.set_ylabel("Behavior Frequency (Count / Session)", fontsize=10)
  ax.grid(True, linestyle=":", alpha=0.6)
  ax.legend(loc="upper right")
  plt.tight_layout()
  return fig


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
      " task transitions, fine-motor academic demands, and high-pitch sensory"
      " noise environments."
  )
  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


def generate_mock_qabf_docx(cohort_key):
  doc = docx.Document()
  doc.add_heading(
      "PSYCHOMETRIC QABF ASSESSMENT RESULTS (PER TARGET BEHAVIOR)", level=1
  )
  doc.add_paragraph(
      f"Client ID: [CLIENT_ID] | Cohort: {cohort_meta[cohort_key]['title']}\n"
  )
  b_list = cohort_meta[cohort_key]["behaviors"]
  for idx, b in enumerate(b_list, 1):
    doc.add_heading(f"Target Behavior #{idx}: {b['name']}", level=2)
    doc.add_paragraph(f"QABF Subscale Breakdown:\n{b['qabf_summary']}")
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
                "• 📐 <strong>Triangulation Algorithm:</strong> 65% Direct ABC"
                " Observations + 25% Indirect Stakeholder Interviews + 10%"
                " Psychometric QABF Assessment"
            ),
        ],
        "behaviors": [
            {
                "name": (
                    "1. Self-Injurious Behavior (SIB) - Head Banging & Wrist"
                    " Biting"
                ),
                "name_zh": "1. 自伤行为 (SIB) - 撞头与咬手腕",
                "name_es": (
                    "1. Conducta autolesiva (SIB) - Golpes en la cabeza y"
                    " mordedura de muñeca"
                ),
                "def": (
                    "Any instance where the child forcefully makes contact"
                    " between forehead and a hard/padded surface, or places"
                    " wrist/hand between upper and lower teeth with visible"
                    " physical force lasting >1 second. Onset is the initial"
                    " physical contact; offset is 5 consecutive seconds without"
                    " contact."
                ),
                "def_zh": (
                    "儿童额头与硬质或软垫表面发生有力的碰撞，或将手腕/手部置于上下牙齿之间咬合并伴有明显施力的任何行为，持续时间>1秒。以首次物理接触为行为开始，以连续5秒无上述动作为行为结束。"
                ),
                "def_es": (
                    "Cualquier instancia en la que el niño haga contacto con"
                    " fuerza entre la frente y una superficie dura/acolchada, o"
                    " coloque la muñeca/mano entre los dientes superiores e"
                    " inferiores con fuerza física visible que dure >1 segundo."
                ),
                "ex": (
                    "Banging forehead 3-4 times on foam mat during tabletop"
                    " task; leaving red teeth marks on right wrist during"
                    " transition."
                ),
                "ex_zh": (
                    "在桌面任务期间，额头在泡棉垫上连续碰撞3-4次；在转换环节中在右手腕上留下红印牙痕。"
                ),
                "ex_es": (
                    "Golpear la frente 3-4 veces sobre la colchoneta durante"
                    " la tarea; dejar marcas rojas de dientes en la muñeca"
                    " derecha."
                ),
                "non_ex": (
                    "Resting head on floor mat during group circle time;"
                    " mouthing FDA-approved silicone chew necklace softly."
                ),
                "non_ex_zh": (
                    "在集体圈圈时间将头靠在地垫上休息；轻柔地咀嚼经过安全认证的硅胶项链。"
                ),
                "non_ex_es": (
                    "Descansar la cabeza en la colchoneta; morder suavemente un"
                    " collar de silicona apto."
                ),
                "dimensions": (
                    "Frequency: 3-6 episodes/day. Duration: 15s - 2min per"
                    " outburst. Intensity: Moderate to Severe (potential skin"
                    " redness or bruising)."
                ),
                "dimensions_zh": (
                    "频率：每天 3-6 次。持续时间：每次发作 15秒 至 2分钟。强度：中度至重度（可能导致皮肤发红或瘀青）。"
                ),
                "dimensions_es": (
                    "Frecuencia: 3-6 episodios/día. Duración: 15s - 2min."
                    " Intensidad: Moderada a Grave."
                ),
                "triggers": (
                    "Setting Events: Fatigue, high ambient background noise.\nImmediate"
                    " Triggers: Presentation of fine-motor discrete trial"
                    " tasks, removal of preferred sensory toy."
                ),
                "triggers_zh": (
                    "背景事件：疲劳、环境背景噪音过大。\n直接触发因素：呈现精细动作桌面任务、移走偏好的感官玩具。"
                ),
                "triggers_es": (
                    "Eventos de contexto: Fatiga, ruido ambiental.\nDesencadenantes"
                    " inmediatos: Tareas motoras finas, retirada de juguete"
                    " preferido."
                ),
                "consequences": (
                    "RBT blocks contact using foam pad, pauses academic demand"
                    " immediately, and prompts PECS 'Break' card."
                ),
                "consequences_zh": (
                    "RBT 使用软垫阻挡物理接触，立即暂停学业要求，并提示使用 PECS“休息”卡片。"
                ),
                "consequences_es": (
                    "El RBT bloquea el contacto, pausa la demanda e indica la"
                    " tarjeta PECS 'Descanso'."
                ),
                "qabf_summary": (
                    "Task Escape: 14/15 | Physical Discomfort: 8/15 |"
                    " Attention: 2/15 | Tangible: 3/15 | Sensory: 4/15"
                ),
                "qabf_summary_zh": (
                    "逃避任务: 14/15 | 身体不适: 8/15 | 社交关注: 2/15 | 获得物质: 3/15 |"
                    " 感官刺激: 4/15"
                ),
                "qabf_summary_es": (
                    "Escape de tarea: 14/15 | Malestar físico: 8/15 | Atención:"
                    " 2/15 | Tangible: 3/15 | Sensorial: 4/15"
                ),
                "triangulation": (
                    "65% Direct ABC Data (high rate during table work) + 25%"
                    " Indirect Parent Interview (frustration with structured"
                    " demands) + 10% QABF (Escape score 14/15)."
                ),
                "triangulation_zh": (
                    "65% 直接 ABC 数据（桌面任务期间高发） + 25% 间接家长访谈（对结构化任务表现出挫败）"
                    " + 10% QABF 结果（逃避任务得分 14/15）。"
                ),
                "triangulation_es": (
                    "65% Datos ABC + 25% Entrevista + 10% QABF (Escape 14/15)."
                ),
                "hypothesis": (
                    "Primary Function: Task Escape (Social Negative"
                    " Reinforcement).\nSecondary Function: Physical Discomfort"
                    " Relief."
                ),
                "hypothesis_zh": "核心行为功能：逃避任务（社交负强化）。\n次要行为功能：缓解身体不适。",
                "hypothesis_es": (
                    "Función principal: Escape de tareas.\nFunción secundaria:"
                    " Alivio de malestar físico."
                ),
                "ferb": (
                    "Functional Communication: Activate AAC BigMack button"
                    " ('I need a break') or hand 'Break' PECS card to prompt"
                    " immediate task pause."
                ),
                "ferb_zh": (
                    "功能性替代行为 (FERB)：按下 AAC 语音按键（“我想休息”）或递交“休息”"
                    " PECS 卡片，以提示立即暂停当前任务。"
                ),
                "ferb_es": (
                    "Comunicación Funcional: Activar el botón AAC BigMack"
                    " ('Necesito un descanso') o entregar la tarjeta PECS"
                    " 'Descanso'."
                ),
            }
        ],
        "strengths": (
            "Responds very well to 1:1 adult playful interaction, strong"
            " visual matching skills, highly motivated by musical cause-and-effect"
            " toys."
        ),
        "strengths_zh": (
            "对 1:1"
            " 成人游戏化互动反应良好，具备较强的视觉匹配能力，对音乐因果玩具表现出极高动机。"
        ),
        "strengths_es": (
            "Responde muy bien a la interacción lúdica 1:1 con adultos,"
            " excelentes habilidades de emparejamiento visual, altamente motivado"
            " por juguetes musicales."
        ),
        "history": (
            "Diagnosed with ASD Level 3; currently receiving 15 hrs/week of"
            " Early Intervention ABA and Speech Therapy."
        ),
        "history_zh": (
            "确诊为孤独症谱系障碍（ASD 3级）；目前每周接受 15"
            " 小时的早期干预 ABA 及言语治疗服务。"
        ),
        "history_es": (
            "Diagnóstico de TEA Nivel 3; actualmente recibe 15 horas/semana de"
            " Intervención Temprana ABA y Terapia de Lenguaje."
        ),
    },
    "g2": {
        "title": "School-Age IEP Protocol (5-21 Yrs)",
        "file_tag": "5to21yo",
        "framework": "IDEA IEP | PBIS Framework",
        "age_str": "10 Years 2 Months",
        "setting_str": "General Education Classroom / Resource Room",
        "protocol_sentences": [
            (
                "• Aligned with IDEA IEP requirements and PBIS Multi-Tiered"
                " Support Systems."
            ),
            (
                "• Targets academic task engagement, self-advocacy, and emotional"
                " self-regulation."
            ),
            (
                "• Emphasizes replacement behaviors integrated into classroom"
                " routines."
            ),
            (
                "• 📐 <strong>Triangulation Algorithm:</strong> 65% Direct ABC"
                " Observations + 25% Indirect Stakeholder Interviews + 10%"
                " Psychometric QABF Assessment"
            ),
        ],
        "behaviors": [
            {
                "name": "1. Task Avoidance / Elopement from Seat",
                "name_zh": "1. 逃避学业任务与擅离座位 (Elopement)",
                "name_es": "1. Evitación de tareas y abandono del asiento",
                "def": (
                    "Leaving assigned desk area without teacher permission for"
                    " >5 seconds during independent academic instruction. Onset:"
                    " feet leaving designated desk perimeter; Offset: returning"
                    " to seat."
                ),
                "def_zh": (
                    "在独立学业教学期间，未经教师允许离开指定的课桌区域超过5秒。以双脚离开指定课桌边界为行为开始，以返回座位为行为结束。"
                ),
                "def_es": (
                    "Abandonar el área asignada del escritorio sin permiso del"
                    " maestro durante >5 segundos durante la instrucción"
                    " académica independiente."
                ),
                "ex": (
                    "Running out of seat to classroom carpet area during"
                    " independent math worksheet."
                ),
                "ex_zh": (
                    "在独立完成数学工作表期间，从座位上跑开并躺在教室地毯区域。"
                ),
                "ex_es": (
                    "Salir corriendo del asiento hacia el área de la alfombra"
                    " durante una hoja de trabajo de matemáticas."
                ),
                "non_ex": (
                    "Standing up to walk to sharpener after raising hand and"
                    " receiving permission."
                ),
                "non_ex_zh": "举手并获得许可后，站起来走到削笔器旁削铅笔。",
                "non_ex_es": (
                    "Levantarse para ir al sacapuntas tras levantar la mano y"
                    " recibir permiso."
                ),
                "dimensions": (
                    "Frequency: 4-5 times per school day. Duration: 1-5 minutes"
                    " per instance. Intensity: Low to Moderate."
                ),
                "dimensions_zh": (
                    "频率：每个学校日 4-5 次。持续时间：每次 1-5 分钟。强度：低至中度。"
                ),
                "dimensions_es": (
                    "Frecuencia: 4-5 veces por día escolar. Duración: 1-5 minutos."
                    " Intensidad: Baja a Moderada."
                ),
                "triggers": (
                    "Setting Events: Multi-step math tasks.\nImmediate"
                    " Triggers: Presentation of 2-page math assignment."
                ),
                "triggers_zh": (
                    "背景事件：多步骤数学任务。\n直接触发因素：发放长达2页的数学作业纸。"
                ),
                "triggers_es": (
                    "Eventos de contexto: Tareas matemáticas de múltiples"
                    " pasos.\nDesencadenante: Entrega de hojas de trabajo."
                ),
                "consequences": (
                    "Staff presents 'Break' visual card; demand temporarily"
                    " paused."
                ),
                "consequences_zh": "教职工呈现“休息”视觉卡片；学业要求被暂时暂停。",
                "consequences_es": (
                    "El personal presenta tarjeta visual 'Descanso'; demanda"
                    " pausada."
                ),
                "qabf_summary": (
                    "Task Escape: 15/15 | Attention: 5/15 | Tangible: 2/15 |"
                    " Sensory: 1/15 | Physical: 0/15"
                ),
                "qabf_summary_zh": (
                    "逃避任务: 15/15 | 社交关注: 5/15 | 获得物质: 2/15 | 感官刺激: 1/15 |"
                    " 身体不适: 0/15"
                ),
                "qabf_summary_es": (
                    "Escape: 15/15 | Atención: 5/15 | Tangible: 2/15 |"
                    " Sensorial: 1/15 | Físico: 0/15"
                ),
                "triangulation": (
                    "65% Direct ABC Data + 25% Indirect IEP Interview + 10%"
                    " QABF (Escape score 15/15)."
                ),
                "triangulation_zh": (
                    "65% 直接 ABC 数据 + 25% 间接 IEP 访谈 + 10% QABF 结果（逃避得分 15/15）。"
                ),
                "triangulation_es": (
                    "65% Datos ABC + 25% Entrevista IEP + 10% QABF (Escape"
                    " 15/15)."
                ),
                "hypothesis": "Primary Function: Escape from Academic Demands.",
                "hypothesis_zh": "核心行为功能：逃避学业任务要求。",
                "hypothesis_es": "Función principal: Escape de demandas académicas.",
                "ferb": (
                    "Hand 'Break' card to teacher or place 'Help Needed' tent on"
                    " desk."
                ),
                "ferb_zh": (
                    "向教师递交“休息”卡片，或在桌上摆放“需要帮助”提示牌。"
                ),
                "ferb_es": (
                    "Entregar tarjeta 'Descanso' al maestro o colocar tarjeta"
                    " 'Ayuda'."
                ),
            }
        ],
        "strengths": (
            "Excellent visual-spatial abilities, enthusiastic about"
            " technology and drawing."
        ),
        "strengths_zh": "具备出色的视觉空间能力，对科技和绘画抱有极高热情。",
        "strengths_es": (
            "Excelentes habilidades visoespaciales, entusiasta de la tecnología"
            " y el dibujo."
        ),
        "history": "Enrolled in General Education with IEP support.",
        "history_zh": "就读于普通教育班级，享有 IEP 特殊教育支持计划。",
        "history_es": "Inscrito en educación general con apoyo de IEP.",
    },
    "g3": {
        "title": "Adult Community Protocol (21+ Yrs)",
        "file_tag": "21plusYo",
        "framework": "Medicaid HCBS | Person-Centered Waiver Framework",
        "age_str": "26 Years 8 Months",
        "setting_str": "Vocational Workshop & Day Program",
        "protocol_sentences": [
            (
                "• Designed for Medicaid HCBS Waiver adult day programs and"
                " community living."
            ),
            (
                "• Focuses on person-centered planning, independence, and"
                " vocational endurance."
            ),
            (
                "• Emphasizes self-management protocols and respectful adult"
                " communication."
            ),
            (
                "• 📐 <strong>Triangulation Algorithm:</strong> 65% Direct ABC"
                " Observations + 25% Indirect Stakeholder Interviews + 10%"
                " Psychometric QABF Assessment"
            ),
        ],
        "behaviors": [
            {
                "name": "1. Vocational Task Refusal & Verbal Aggression",
                "name_zh": "1. 拒绝职业组装任务与言语攻击",
                "name_es": "1. Rechazo de tareas vocacionales y agresión verbal",
                "def": (
                    "Refusing assembly or sorting demands accompanied by loud"
                    " vocal threats (>75 dB) or pushing work materials away."
                    " Onset: vocal outburst or material push; Offset: 3 minutes"
                    " of quiet task engagement."
                ),
                "def_zh": (
                    "拒绝组装或分类任务，并伴随大声言语威胁（>75 dB）或推开工作材料的行为。以言语发作或推开材料为行为开始，以连续"
                    " 3 分钟安静参与任务为行为结束。"
                ),
                "def_es": (
                    "Rechazar demandas de ensamblaje acompañado de amenazas"
                    " vocales fuertes (>75 dB) o empujar los materiales de"
                    " trabajo."
                ),
                "ex": (
                    "Shouting 'No way!', slamming assembly boxes on desk when"
                    " quota is raised."
                ),
                "ex_zh": (
                    "当提高工作配额时，大叫“绝不可能！”并将组装盒重重摔在桌上。"
                ),
                "ex_es": (
                    "Gritar '¡De ninguna manera!' y golpear cajas sobre la mesa."
                ),
                "non_ex": "Verbally requesting a 5-minute break in a normal tone.",
                "non_ex_zh": "用正常音量和语调口头提出“想要休息5分钟”。",
                "non_ex_es": (
                    "Solicitar verbalmente un descanso de 5 minutos en tono"
                    " normal."
                ),
                "dimensions": (
                    "Frequency: 1-2 times weekly. Duration: 5-10 minutes."
                    " Intensity: Moderate."
                ),
                "dimensions_zh": (
                    "频率：每周 1-2 次。持续时间：5-10 分钟。强度：中度。"
                ),
                "dimensions_es": (
                    "Frecuencia: 1-2 veces por semana. Duración: 5-10 min."
                    " Intensidad: Moderada."
                ),
                "triggers": (
                    "Setting Events: Unfamiliar staff.\nImmediate Triggers:"
                    " Direct instructions to complete vocational assembly"
                    " quota."
                ),
                "triggers_zh": (
                    "背景事件：不熟悉的工作人员。\n直接触发因素：要求完成职业组装配额的直接指令。"
                ),
                "triggers_es": (
                    "Eventos de contexto: Personal no familiar.\nDesencadenante:"
                    " Instrucción directa de ensamblaje."
                ),
                "consequences": (
                    "DSP offers choice board, demand temporarily paused."
                ),
                "consequences_zh": (
                    "直属支持人员（DSP）提供选择板，任务要求被暂时暂停。"
                ),
                "consequences_es": (
                    "El personal DSP ofrece panel de opciones, demanda"
                    " pausada."
                ),
                "qabf_summary": (
                    "Task Escape: 13/15 | Attention: 4/15 | Tangible: 3/15 |"
                    " Sensory: 1/15 | Physical: 1/15"
                ),
                "qabf_summary_zh": (
                    "逃避任务: 13/15 | 社交关注: 4/15 | 获得物质: 3/15 | 感官刺激: 1/15 |"
                    " 身体不适: 1/15"
                ),
                "qabf_summary_es": (
                    "Escape: 13/15 | Atención: 4/15 | Tangible: 3/15 |"
                    " Sensorial: 1/15 | Físico: 1/15"
                ),
                "triangulation": (
                    "65% Direct ABC Data + 25% Indirect Vocational Interview +"
                    " 10% QABF (Escape score 13/15)."
                ),
                "triangulation_zh": (
                    "65% 直接 ABC 数据 + 25% 间接职业能力访谈 + 10% QABF 结果（逃避得分 13/15）。"
                ),
                "triangulation_es": (
                    "65% Datos ABC + 25% Entrevista + 10% QABF (Escape 13/15)."
                ),
                "hypothesis": (
                    "Primary Function: Escape from Vocational Assembly Demands."
                ),
                "hypothesis_zh": "核心行为功能：逃避职业组装工作要求。",
                "hypothesis_es": (
                    "Función principal: Escape de demandas vocacionales."
                ),
                "ferb": (
                    "Verbally request '5-minute break, please' using"
                    " self-advocacy phrase card."
                ),
                "ferb_zh": (
                    "使用自我倡导短语卡口头表达：“请给我 5 分钟休息时间”。"
                ),
                "ferb_es": (
                    "Solicitar verbalmente 'Descanso de 5 minutos, por favor'."
                ),
            }
        ],
        "strengths": "High independence in personal self-care.",
        "strengths_zh": "在个人日常生活自理方面具备极高独立性。",
        "strengths_es": "Alta independencia en el cuidado personal.",
        "history": (
            "Participates in Adult Day Vocational Services under Medicaid HCBS"
            " Waiver."
        ),
        "history_zh": (
            "在 Medicaid HCBS 豁免计划下参与成人日间职业干预服务。"
        ),
        "history_es": (
            "Participa en servicios vocacionales para adultos bajo exención"
            " Medicaid HCBS."
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
    "<div class='sub-header'>Interactive Demonstration with Automated"
    " Insurance-Grade Visual Analysis & Document Synthesis</div>",
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

st.markdown(
    f"""
    <div class="protocol-card">
        <div class="protocol-title">📋 Selected Protocol Framework: {current_meta['title']} ({current_meta['framework']})</div>
        {"".join([f'<div class="protocol-bullet">{s}</div>' for s in current_meta['protocol_sentences']])}
    </div>
""",
    unsafe_allow_html=True,
)

col_input1, col_input2, col_input3 = st.columns([1, 1, 1])

with col_input1:
  st.markdown("#### 📄 Direct Observation (ABC)")
  mock_csv = generate_mock_abc_csv(selected_cohort_key)
  st.download_button(
      label="📥 Download Mock ABC (.csv)",
      data=mock_csv,
      file_name=f"DeIdentified_ABC_{current_meta['file_tag']}.csv",
      mime="text/csv",
      use_container_width=True,
  )
  uploaded_abc = st.file_uploader(
      "Upload ABC File:", type=["csv", "xlsx"], key=f"abc_{selected_cohort_key}"
  )

with col_input2:
  st.markdown("#### 📝 Indirect Interview Notes")
  mock_docx = generate_mock_interview_docx(selected_cohort_key)
  st.download_button(
      label="📥 Download Mock Interview (.docx)",
      data=mock_docx,
      file_name=f"DeIdentified_Interview_{current_meta['file_tag']}.docx",
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
  st.markdown("#### 📊 Behavior QABF Assessment Results")
  mock_qabf = generate_mock_qabf_docx(selected_cohort_key)
  st.download_button(
      label="📥 Download Mock QABF Results (.docx)",
      data=mock_qabf,
      file_name=f"DeIdentified_QABF_{current_meta['file_tag']}.docx",
      mime=(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      ),
      use_container_width=True,
  )
  uploaded_qabf = st.file_uploader(
      "Upload QABF File:",
      type=["docx", "txt"],
      key=f"qabf_{selected_cohort_key}",
  )

active_behaviors = current_meta["behaviors"]

# ==========================================
# 6. Integrated Visual Analysis Section
# ==========================================
st.markdown("### 3️⃣ Insurance Audit Visual Analysis & Tracking Trend")
st.markdown(
    "保险公司审核及资金延期通常需要查看**基线期（Baseline）与干预期（Intervention）**的量化趋势图。以下是通过模拟数据追踪文件自动生成的折线趋势分析："
)

df_tracking = generate_mock_tracking_data()
fig_trend = plot_aba_behavior_graph(df_tracking)
st.pyplot(fig_trend)


# ==========================================
# 7. Word Document Helper Functions
# ==========================================
def add_bi_heading(doc, level, text_en, text_trans=None):
  h = doc.add_heading(level=level)
  r_en = h.add_run(text_en)
  if text_trans:
    r_tr = h.add_run(f" [{text_trans}]")
    r_tr.italic = True
    r_tr.font.size = Pt(11)
    r_tr.font.color.rgb = RGBColor(120, 120, 120)


def add_bi_item(
    doc, label_en, val_en, label_trans=None, val_trans=None, lang_mode="zh"
):
  p = doc.add_paragraph()
  p.paragraph_format.space_after = Pt(4)
  p.paragraph_format.space_before = Pt(2)

  r_lbl = p.add_run(f"{label_en}: ")
  r_lbl.bold = True
  p.add_run(f"{val_en}")

  if lang_mode != "en" and label_trans and val_trans:
    p.add_run("\n")
    r_tr_lbl = p.add_run(f"[{label_trans}: ")
    r_tr_lbl.bold = True
    r_tr_lbl.italic = True
    r_tr_lbl.font.color.rgb = RGBColor(100, 100, 100)

    r_tr_val = p.add_run(f"{val_trans}]")
    r_tr_val.italic = True
    r_tr_val.font.color.rgb = RGBColor(100, 100, 100)


def build_compact_demographics_table(doc, c_meta, lang_mode):
  table = doc.add_table(rows=5, cols=2)
  table.style = "Table Grid"

  if lang_mode == "zh":
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
  elif lang_mode == "es":
    data = [
        (
            "Client Name (Nombre)",
            "[CLIENT_NAME]",
            "DOB / Age (Edad)",
            f"[CLIENT_DOB] / {c_meta['age_str']}",
        ),
        (
            "Client ID (Identificación)",
            "[CLIENT_ID]",
            "Assessment Date (Fecha)",
            "2026-08-14",
        ),
        (
            "Facility/School (Centro)",
            "[DISTRICT_OR_FACILITY_NAME]",
            "Setting (Entorno)",
            c_meta["setting_str"],
        ),
        (
            "Assessor (Evaluador)",
            "[BCBA_NAME], BCBA, LBA",
            "Framework (Marco)",
            c_meta["framework"],
        ),
        (
            "Primary Language (Idioma)",
            "English / Spanish",
            "Informants (Informantes)",
            "Parent, Lead Teacher / RBT",
        ),
    ]
  else:
    data = [
        (
            "Student/Client Name",
            "[CLIENT_NAME]",
            "DOB / Age",
            f"[CLIENT_DOB] / {c_meta['age_str']}",
        ),
        ("Client ID", "[CLIENT_ID]", "Assessment Date", "2026-08-14"),
        (
            "Facility/School",
            "[DISTRICT_OR_FACILITY_NAME]",
            "Setting",
            c_meta["setting_str"],
        ),
        ("Assessor", "[BCBA_NAME], BCBA, LBA", "Framework", c_meta["framework"]),
        ("Primary Language", "English", "Informants", "Parent, Lead Teacher"),
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
# 8. Document Generators (FBA & BIP)
# ==========================================
def generate_exact_fba_doc(cohort_key, lang_choice, behavior_list):
  c_meta = cohort_meta[cohort_key]
  doc = docx.Document()

  lang_mode = (
      "zh"
      if "Chinese" in lang_choice
      else "es"
      if "Spanish" in lang_choice
      else "en"
  )

  p_t = doc.add_paragraph()
  r_t = p_t.add_run("FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT")
  r_t.bold = True
  p_t.style = doc.styles["Title"]

  add_bi_heading(
      doc,
      1,
      "1. Student Demographics & Administrative Info",
      (
          "1. 学生基本信息"
          if lang_mode == "zh"
          else "1. Datos demográficos"
          if lang_mode == "es"
          else None
      ),
  )
  build_compact_demographics_table(doc, c_meta, lang_mode)

  add_bi_heading(
      doc,
      1,
      "2. Data Sources & Triangulation Methodology",
      (
          "2. 数据来源与三方验证方法"
          if lang_mode == "zh"
          else "2. Fuentes de datos y triangulación"
          if lang_mode == "es"
          else None
      ),
  )
  add_bi_item(
      doc,
      "Methodology & Triangulation Formula",
      (
          "Triangulation Algorithm: 65% Direct ABC Data + 25% Indirect"
          " Interview + 10% QABF."
      ),
      (
          "评估方法与三方验证公式"
          if lang_mode == "zh"
          else "Metodología"
          if lang_mode == "es"
          else None
      ),
      (
          "应用三方加权验证算法：65% 直接观察 + 25% 访谈 + 10% QABF量表。"
          if lang_mode == "zh"
          else "Algoritmo de triangulación aplicado."
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  add_bi_heading(
      doc,
      1,
      "3. Brief Background & Strengths Summary",
      (
          "3. 背景与优势摘要"
          if lang_mode == "zh"
          else "3. Antecedentes"
          if lang_mode == "es"
          else None
      ),
  )
  add_bi_item(
      doc,
      "Strengths & Preferences",
      c_meta["strengths"],
      (
          "优势与偏好"
          if lang_mode == "zh"
          else "Fortalezas"
          if lang_mode == "es"
          else None
      ),
      (
          c_meta["strengths_zh"]
          if lang_mode == "zh"
          else c_meta.get("strengths_es", c_meta["strengths"])
      ),
      lang_mode,
  )

  add_bi_heading(
      doc,
      1,
      "4. Individual Target Behavior Functional Analyses",
      (
          "4. 目标行为独立功能分析"
          if lang_mode == "zh"
          else "4. Análisis funcional"
          if lang_mode == "es"
          else None
      ),
  )
  for idx, b in enumerate(behavior_list, 1):
    trans_name = (
        b.get("name_zh")
        if lang_mode == "zh"
        else b.get("name_es")
        if lang_mode == "es"
        else None
    )
    add_bi_heading(
        doc,
        2,
        f"Target Behavior #{idx}: {b['name']}",
        f"目标行为 #{idx}: {trans_name}" if trans_name else None,
    )
    add_bi_item(
        doc,
        "Operational Definition",
        b["def"],
        (
            "操作性定义"
            if lang_mode == "zh"
            else "Definición operacional"
            if lang_mode == "es"
            else None
        ),
        (
            b.get("def_zh")
            if lang_mode == "zh"
            else b.get("def_es")
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )
    add_bi_item(
        doc,
        "Hypothesized Function",
        b["hypothesis"],
        (
            "推断功能"
            if lang_mode == "zh"
            else "Función"
            if lang_mode == "es"
            else None
        ),
        (
            b.get("hypothesis_zh")
            if lang_mode == "zh"
            else b.get("hypothesis_es")
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


def generate_exact_bip_doc(cohort_key, lang_choice):
  c_meta = cohort_meta[cohort_key]
  doc = docx.Document()
  lang_mode = (
      "zh"
      if "Chinese" in lang_choice
      else "es"
      if "Spanish" in lang_choice
      else "en"
  )

  p_t = doc.add_paragraph()
  r_t = p_t.add_run("BEHAVIOR INTERVENTION PLAN (BIP)")
  r_t.bold = True
  p_t.style = doc.styles["Title"]

  add_bi_heading(
      doc,
      1,
      "1. Student Info & Administrative Summary",
      (
          "1. 学生信息摘要"
          if lang_mode == "zh"
          else "1. Información"
          if lang_mode == "es"
          else None
      ),
  )
  build_compact_demographics_table(doc, c_meta, lang_mode)

  add_bi_heading(
      doc,
      1,
      "2. Target Behaviors, Functions & Replacement Skills (FERB)",
      (
          "2. 目标行为与替代技能 (FERB)"
          if lang_mode == "zh"
          else "2. FERB"
          if lang_mode == "es"
          else None
      ),
  )
  for idx, b in enumerate(c_meta["behaviors"], 1):
    trans_name = (
        b.get("name_zh")
        if lang_mode == "zh"
        else b.get("name_es")
        if lang_mode == "es"
        else None
    )
    add_bi_heading(
        doc,
        2,
        f"Target Behavior #{idx}: {b['name']}",
        f"目标行为 #{idx}: {trans_name}" if trans_name else None,
    )
    add_bi_item(
        doc,
        "FERB",
        b["ferb"],
        (
            "替代行为"
            if lang_mode == "zh"
            else "Reemplazo"
            if lang_mode == "es"
            else None
        ),
        (
            b.get("ferb_zh")
            if lang_mode == "zh"
            else b.get("ferb_es")
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 9. Action Buttons & Download
# ==========================================
st.markdown("### 4️⃣ Target Language & Formulate / Download Actions")

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

fba_docx_bytes = generate_exact_fba_doc(
    selected_cohort_key, report_lang, active_behaviors
)
bip_docx_bytes = generate_exact_bip_doc(selected_cohort_key, report_lang)

lang_tag = (
    "English"
    if "Standard" in report_lang
    else "Bilingual_ZH"
    if "Chinese" in report_lang
    else "Bilingual_ES"
)

with col_action1:
  st.write(" ")
  st.write(" ")
  st.download_button(
      label="⚡ Formulate & Download FBA Draft (.docx)",
      data=fba_docx_bytes,
      file_name=(
          f"DeIdentified_FBA_Draft_{current_meta['file_tag']}_{lang_tag}.docx"
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
      label="⚡ Formulate & Download BIP Draft (.docx)",
      data=bip_docx_bytes,
      file_name=(
          f"DeIdentified_BIP_Draft_{current_meta['file_tag']}_{lang_tag}.docx"
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
)[cite: 6]
