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
    .algo-highlight { font-weight: 700; color: #C0392B; background-color: #FDEDEC; padding: 2px 6px; border-radius: 4px; }
    
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
              "Date_Time": "2026-07-20 09:15",
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
              "Date_Time": "2026-07-27 10:30",
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
              "Date_Time": "2026-08-03 11:00",
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
              "Date_Time": "2026-08-10 15:20",
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
              "Date_Time": "2026-07-20 09:30",
              "Setting": "Gen-Ed Classroom",
              "Antecedent": "Teacher presented 2-page math worksheet",
              "Behavior": "Screamed 'I won't do it!', pushed desk away",
              "Consequence": (
                  "Staff presented 'Break' visual card; demand paused"
              ),
          },
          {
              "Date_Time": "2026-07-27 11:00",
              "Setting": "Science Lab",
              "Antecedent": "Group presentation requirement announced",
              "Behavior": "Left desk area and hid under back table for 4 minutes",
              "Consequence": "Peer partner retrieved folder; teacher gave prompt",
          },
          {
              "Date_Time": "2026-08-03 13:15",
              "Setting": "School Cafeteria",
              "Antecedent": "Loud ambient cafeteria chatter and clattering trays",
              "Behavior": (
                  "Covered ears, vocalized high pitch, and bolted toward exit"
              ),
              "Consequence": (
                  "Aide followed with hall pass, escorted to quiet library"
                  " corner"
              ),
          },
          {
              "Date_Time": "2026-08-10 14:30",
              "Setting": "Computer Lab",
              "Antecedent": "Login failure on educational software page",
              "Behavior": (
                  "Struck keyboard keys forcefully and shoved monitor back"
              ),
              "Consequence": "IT support reset password; teacher provided help",
          },
      ],
      "g3": [
          {
              "Date_Time": "2026-07-20 09:00",
              "Setting": "Residential Home",
              "Antecedent": "New staff worker introduced schedule change",
              "Behavior": "Swept dishes off table, shouted threats",
              "Consequence": (
                  "DSP stepped in, offered visual choice board, demand paused"
              ),
          },
          {
              "Date_Time": "2026-07-27 10:15",
              "Setting": "Vocational Workshop",
              "Antecedent": "Supervisor increased box assembly quota by 20 units",
              "Behavior": "Shouted 'This is garbage!', slammed parts into bin",
              "Consequence": "Job coach offered 10-minute walk break",
          },
          {
              "Date_Time": "2026-08-03 13:30",
              "Setting": "Community Living Room",
              "Antecedent": (
                  "Peer requested remote control to change television channel"
              ),
              "Behavior": "Grabbed cushion, threw it across room, cursed loudly",
              "Consequence": "Staff intervened, redirected peer to another room",
          },
          {
              "Date_Time": "2026-08-10 11:00",
              "Setting": "Kitchen Prep Area",
              "Antecedent": "Staff asked client to wash cooking pans immediately",
              "Behavior": "Crossed arms, turned back, refused to touch sponge",
              "Consequence": "Staff gave 5-minute transition timer and space",
          },
      ],
  }
  df = pd.DataFrame(datasets.get(cohort_key, datasets["g1"]))
  return df.to_csv(index=False).encode("utf-8")


def generate_mock_tracking_csv(cohort_key):
  data = {
      "Week": [
          "Week 1 (Jul 20)",
          "Week 2 (Jul 27)",
          "Week 3 (Aug 03)",
          "Week 4 (Aug 10)",
      ],
      "Target_Behavior_1_Freq": [18, 14, 9, 5],
      "Target_Behavior_2_Freq": [12, 10, 7, 4],
      "Target_Behavior_3_Freq": [15, 11, 8, 3],
  }
  df = pd.DataFrame(data)
  return df.to_csv(index=False).encode("utf-8")


def generate_behavior_tracking_chart(cohort_key, behavior_index):
  fig, ax = plt.subplots(figsize=(6, 2.5))
  weeks = [
      "Week 1 (Jul 20)",
      "Week 2 (Jul 27)",
      "Week 3 (Aug 03)",
      "Week 4 (Aug 10)",
  ]

  if behavior_index == 0:
    freqs = [18, 14, 9, 5]
    color_val = "#1F4E78"
    title_str = "Target Behavior #1: 4-Week Frequency Trend"
  elif behavior_index == 1:
    freqs = [12, 10, 7, 4]
    color_val = "#C0392B"
    title_str = "Target Behavior #2: 4-Week Frequency Trend"
  else:
    freqs = [15, 11, 8, 3]
    color_val = "#27AE60"
    title_str = "Target Behavior #3: 4-Week Frequency Trend"

  ax.plot(
      weeks, freqs, marker="o", color=color_val, linewidth=2, markersize=6
  )
  ax.set_title(
      title_str,
      fontsize=9,
      fontweight="bold",
      color=color_val,
  )
  ax.set_xlabel("4-Week Observation Period", fontsize=8)
  ax.set_ylabel("Total Weekly Episodes", fontsize=8)
  ax.grid(True, linestyle="--", alpha=0.5)
  plt.xticks(rotation=15, fontsize=8)
  plt.tight_layout()
  buf = io.BytesIO()
  fig.savefig(buf, format="png", dpi=150)
  buf.seek(0)
  plt.close(fig)
  return buf


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
  doc.add_heading("1. Interview Note Segment 1 (Parent Perspective)", level=2)
  doc.add_paragraph(
      "Summary: Stakeholders report elevated rates of target behaviors during"
      " task transitions across a 4-week observation window, fine-motor"
      " academic demands, and high-pitch sensory noise environments."
  )
  doc.add_heading(
      "2. Interview Note Segment 2 (Educator / RBT Perspective)", level=2
  )
  doc.add_paragraph(
      "Summary: Staff observe functional escape behaviors when demands are"
      " presented rapidly without priming or visual schedules."
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
                " routines across 4 weeks."
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
                    "Frequency: 3-6 episodes/day (tracked over 4 weeks). Duration:"
                    " 15s - 2min per outburst. Intensity: Moderate to Severe"
                    " (potential skin redness or bruising)."
                ),
                "dimensions_zh": (
                    "频率：每天 3-6 次（4周追踪）。持续时间：每次发作 15秒 至 2分钟。强度：中度至重度（可能导致皮肤发红或瘀青）。"
                ),
                "dimensions_es": (
                    "Frecuencia: 3-6 episodios/día (4 semanas). Duración: 15s -"
                    " 2min. Intensidad: Moderada a Grave."
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
                    "65% Direct ABC Data (high rate during table work across 4"
                    " weeks) + 25% Indirect Parent Interview (frustration with"
                    " structured demands) + 10% QABF (Escape score 14/15)."
                ),
                "triangulation_zh": (
                    "65% 直接 ABC 数据（4周内桌面任务期间高发） + 25% 间接家长访谈（对结构化任务表现出挫败）"
                    " + 10% QABF 结果（逃避任务得分 14/15）。"
                ),
                "triangulation_es": (
                    "65% Datos ABC (4 semanas) + 25% Entrevista + 10% QABF"
                    " (Escape 14/15)."
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
            },
            {
                "name": "2. Sensory Vocal Distress & Face-Slapping",
                "name_zh": "2. 感官情绪性尖叫与拍打面部",
                "name_es": "2. Distrés vocal sensorial y bofetadas en la cara",
                "def": (
                    "Vocal screaming exceeding normal conversational volume (>80"
                    " dB) lasting >2 seconds, occurring simultaneously with"
                    " open-palm striking of own cheeks or arms. Onset: initial"
                    " vocalization/strike; Offset: 10 consecutive seconds of"
                    " calm without striking."
                ),
                "def_zh": (
                    "尖叫声超过正常交谈音量（>80 dB）且持续时间>2秒，同时伴随用张开的手掌拍打自己脸颊或手臂的行为。以初始发声/拍打为行为开始，以连续10秒保持平静无拍打动作为行为结束。"
                ),
                "def_es": (
                    "Gritos vocales que superan el volumen conversacional normal"
                    " (>80 dB) durante >2 segundos, junto con golpes con la"
                    " palma abierta en las mejillas o brazos."
                ),
                "ex": (
                    "Loud screaming and striking cheeks 3 times when loud"
                    " kitchen appliance starts in adjacent room."
                ),
                "ex_zh": "当隔壁房间开启高分贝厨房电器时，大声尖叫并连续拍打脸颊 3 次。",
                "ex_es": (
                    "Gritos fuertes y 3 golpes en las mejillas al encenderse un"
                    " electrodoméstico ruidoso."
                ),
                "non_ex": (
                    "Excited vocal shouting during outdoor playground play;"
                    " tapping cheeks rhythmically during music group time."
                ),
                "non_ex_zh": (
                    "在户外操场游玩时兴奋地高声欢呼；在音乐课上跟随节奏轻拍脸颊。"
                ),
                "non_ex_es": (
                    "Gritos de emoción en el parque; golpear las mejillas al"
                    " ritmo de la música."
                ),
                "dimensions": (
                    "Frequency: 2-4 episodes/day (4-week trend). Duration: 30s"
                    " - 3min. Intensity: Moderate."
                ),
                "dimensions_zh": (
                    "频率：每天 2-4 次（4周趋势）。持续时间：30秒 至 3分钟。强度：中度。"
                ),
                "dimensions_es": (
                    "Frecuencia: 2-4 episodios/día (4 semanas). Duración: 30s -"
                    " 3min. Intensidad: Moderada."
                ),
                "triggers": (
                    "Setting Events: Overstimulating ambient noise, sudden"
                    " schedule changes.\nImmediate Triggers: Unexpected high-pitch"
                    " sounds, removal of preferred juice cup."
                ),
                "triggers_zh": (
                    "背景事件：环境噪音过度刺激、突发日程改变。\n直接触发因素：意料之外的高频噪音、偏好的果汁杯被移走。"
                ),
                "triggers_es": (
                    "Eventos de contexto: Ruido ambiental excesivo.\nDesencadenantes:"
                    " Sonidos agudos inesperados."
                ),
                "consequences": (
                    "Staff offers noise-canceling headphones and redirects to"
                    " sensory chew tool."
                ),
                "consequences_zh": "工作人员提供降噪耳机，并重新引导至口部感官咀嚼工具。",
                "consequences_es": (
                    "El personal ofrece auriculares de cancelación de ruido y"
                    " redirige a un mordedor sensorial."
                ),
                "qabf_summary": (
                    "Sensory/Automatic: 13/15 | Task Escape: 11/15 | Attention:"
                    " 3/15 | Tangible: 2/15 | Physical: 1/15"
                ),
                "qabf_summary_zh": (
                    "感官刺激/自动强化: 13/15 | 逃避任务: 11/15 | 社交关注: 3/15 |"
                    " 获得物质: 2/15 | 身体不适: 1/15"
                ),
                "qabf_summary_es": (
                    "Sensorial: 13/15 | Escape: 11/15 | Atención: 3/15 |"
                    " Tangible: 2/15 | Físico: 1/15"
                ),
                "triangulation": (
                    "65% Direct ABC Data (appliance noise trigger over 4"
                    " weeks) + 25% Indirect RBT Interview (auditory aversion) +"
                    " 10% QABF (Sensory score 13/15)."
                ),
                "triangulation_zh": (
                    "65% 直接 ABC 数据（4周内电器噪音触发） + 25% 间接 RBT 访谈（听觉敏感过度） + 10%"
                    " QABF 结果（感官得分 13/15）。"
                ),
                "triangulation_es": (
                    "65% Datos ABC (4 semanas) + 25% Entrevista RBT + 10% QABF"
                    " (Sensorial 13/15)."
                ),
                "hypothesis": (
                    "Primary Function: Escape / Regulation of Auditory"
                    " Overstimulation (Sensory Automatic Reinforcement)."
                ),
                "hypothesis_zh": (
                    "核心行为功能：逃避/调节听觉过度刺激（感官自动强化）。"
                ),
                "hypothesis_es": (
                    "Función principal: Escape / Regulación de sobreestimulación"
                    " auditiva."
                ),
                "ferb": (
                    "Self-Regulation / Request Strategy: Point to 'Headphones'"
                    " visual symbol or independently retrieve noise-canceling"
                    " headphones from designated sensory bin."
                ),
                "ferb_zh": (
                    "功能性替代行为 (FERB)：指认“耳机”视觉标识，或自行从指定的感官箱中取出降噪耳机戴上。"
                ),
                "ferb_es": (
                    "Estrategia de autorregulación: Señalar el símbolo"
                    " 'Auriculares' o tomarlos de la caja sensorial."
                ),
            },
            {
                "name": "3. Floor Dropping & Pica Attempts",
                "name_zh": "3. 突然倒地与异食倾向 (Pica)",
                "name_es": "3. Tirarse al suelo e intentos de pica",
                "def": (
                    "Unprompted sudden collapse of body onto floor from standing"
                    " or seated position, accompanied by attempts to place"
                    " non-food items (e.g., sand, dirt, small plastic objects)"
                    " into mouth. Onset: body contacting floor; Offset:"
                    " standing up for >5 seconds."
                ),
                "def_zh": (
                    "在站立或坐姿状态下，未经提示突然将身体倒在地上，并伴随试图将非食物物品（如沙子、泥土、塑料小零件）放入口中的行为。以身体接触地面为行为开始，以重新站立保持>5秒为行为结束。"
                ),
                "def_es": (
                    "Caída repentina e imprevista del cuerpo al suelo desde una"
                    " posición de pie o sentada, acompañada de intentos de"
                    " llevarse a la boca objetos no alimentarios (p. ej.,"
                    " arena, tierra, objetos pequeños de plástico). Inicio:"
                    " contacto del cuerpo con el suelo; Fin: permanecer de pie"
                    " durante >5 segundos."
                ),
                "ex": (
                    "Dropping to playground sand and placing a handful of sand"
                    " near mouth when prompted to transition inside."
                ),
                "ex_zh": (
                    "当被提示转换进室内时，突然倒在操场沙坑里并将一把沙子拿到嘴边。"
                ),
                "ex_es": (
                    "Tirarse a la arena del parque y llevarse un puñado cerca de"
                    " la boca al recibir la indicación de entrar."
                ),
                "non_ex": (
                    "Lying down on rest mat during scheduled nap time; placing"
                    " silicone chew toy into mouth."
                ),
                "non_ex_zh": (
                    "在计划的午休时间躺在休息垫上；将硅胶咀嚼玩具放入口中。"
                ),
                "non_ex_es": (
                    "Acostarse en la colchoneta durante la siesta; ponerse el"
                    " mordedor de silicona en la boca."
                ),
                "dimensions": (
                    "Frequency: 1-3 episodes/day (4-week tracking). Duration:"
                    " 1-5min. Intensity: Moderate."
                ),
                "dimensions_zh": (
                    "频率：每天 1-3 次（4周追踪）。持续时间：1 至 5分钟。强度：中度。"
                ),
                "dimensions_es": (
                    "Frecuencia: 1-3 episodios/día (4 semanas). Duración: 1-5min."
                    " Intensidad: Moderada."
                ),
                "triggers": (
                    "Setting Events: Unassigned downtime, transition from"
                    " outdoor to indoor.\nImmediate Triggers: Direct"
                    " instruction to clean up preferred toys."
                ),
                "triggers_zh": (
                    "背景事件：无安排的空闲时间、从户外到室内的活动转换。\n直接触发因素：要求收拾偏好玩具的直接指令。"
                ),
                "triggers_es": (
                    "Eventos de contexto: Tiempo libre sin estructurar,"
                    " transición exterior/interior.\nDesencadenante: Orden de"
                    " recoger los juguetes."
                ),
                "consequences": (
                    "RBT blocks hand-to-mouth trajectory, provides oral chew"
                    " device, and delays transition demand."
                ),
                "consequences_zh": (
                    "RBT 阻挡手移向嘴部的轨迹，提供口部咀嚼器，并延缓转换要求。"
                ),
                "consequences_es": (
                    "El RBT bloquea el trayecto mano-boca, ofrece mordedor"
                    " oral y retrasa la demanda."
                ),
                "qabf_summary": (
                    "Task Escape: 12/15 | Tangible Access: 10/15 | Attention:"
                    " 4/15 | Sensory: 3/15 | Physical: 1/15"
                ),
                "qabf_summary_zh": (
                    "逃避任务: 12/15 | 获得物质: 10/15 | 社交关注: 4/15 | 感官刺激:"
                    " 3/15 | 身体不适: 1/15"
                ),
                "qabf_summary_es": (
                    "Escape: 12/15 | Acceso a Tangible: 10/15 | Atención: 4/15"
                    " | Sensorial: 3/15 | Físico: 1/15"
                ),
                "triangulation": (
                    "65% Direct ABC Data (playground cleanup trigger over 4"
                    " weeks) + 25% Indirect Teacher Interview (transition delay)"
                    " + 10% QABF (Escape score 12/15)."
                ),
                "triangulation_zh": (
                    "65% 直接 ABC 数据（4周内操场清理触发） + 25% 间接教师访谈（转换延缓） + 10%"
                    " QABF 结果（逃避得分 12/15）。"
                ),
                "triangulation_es": (
                    "65% Datos ABC (4 semanas) + 25% Entrevista + 10% QABF"
                    " (Escape 12/15)."
                ),
                "hypothesis": (
                    "Primary Function: Transition Escape & Delay.\nSecondary"
                    " Function: Access to Outdoor Tangible Play."
                ),
                "hypothesis_zh": (
                    "核心行为功能：逃避与延缓活动转换。\n次要行为功能：继续获取户外游戏物品。"
                ),
                "hypothesis_es": (
                    "Función principal: Escape/retraso de transición.\nFunción"
                    " secundaria: Acceso a juego."
                ),
                "ferb": (
                    "Transition Delay AAC: Exchange '1 More Minute' visual"
                    " card or press AAC button to request additional play time"
                    " prior to cleanup."
                ),
                "ferb_zh": (
                    "功能性替代行为 (FERB)：递交“再玩1分钟”视觉卡片或按下 AAC"
                    " 沟通按键，在收拾前请求额外的游戏时间。"
                ),
                "ferb_es": (
                    "Comunicación Funcional: Entregar tarjeta '1 minuto más' o"
                    " presionar el botón AAC."
                ),
            },
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
                " Support Systems across 4 weeks."
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
                    "Frequency: 4-5 times per school day (tracked over 4"
                    " weeks). Duration: 1-5 minutes per instance. Intensity:"
                    " Low to Moderate."
                ),
                "dimensions_zh": (
                    "频率：每个学校日 4-5 次（4周追踪）。持续时间：每次 1-5 分钟。强度：低至中度。"
                ),
                "dimensions_es": (
                    "Frecuencia: 4-5 veces por día escolar (4 semanas)."
                    " Duración: 1-5 minutos. Intensidad: Baja a Moderada."
                ),
                "triggers": (
                    "Setting Events: Multi-step math tasks.\nImmediate Triggers:"
                    " Presentation of 2-page math assignment."
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
                    "逃避任务: 15/15 | 社交关注: 5/15 | 获得物质: 2/15 | 感官刺激:"
                    " 1/15 | 身体不适: 0/15"
                ),
                "qabf_summary_es": (
                    "Escape: 15/15 | Atención: 5/15 | Tangible: 2/15 |"
                    " Sensorial: 1/15 | Físico: 0/15"
                ),
                "triangulation": (
                    "65% Direct ABC Data (4 weeks) + 25% Indirect IEP Interview"
                    " + 10% QABF (Escape score 15/15)."
                ),
                "triangulation_zh": (
                    "65% 直接 ABC 数据（4周） + 25% 间接 IEP 访谈 + 10% QABF 结果（逃避得分 15/15）。"
                ),
                "triangulation_es": (
                    "65% Datos ABC (4 semanas) + 25% Entrevista IEP + 10% QABF"
                    " (Escape 15/15)."
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
            },
            {
                "name": "2. Group Presentation Avoidance & Hiding",
                "name_zh": "2. 逃避小组展示与躲藏行为",
                "name_es": "2. Evitación de presentaciones grupales y ocultamiento",
                "def": (
                    "Leaving assigned group space or crawling under furniture"
                    " when public speaking or group presentation is requested."
                    " Onset: body moving away from group; Offset: returning after"
                    " prompt."
                ),
                "def_zh": (
                    "当被要求公开演讲或小组展示时，离开指定的小组空间或爬到家具下方。以身体离开小组为行为开始，以经提示后返回为行为结束。"
                ),
                "def_es": (
                    "Abandonar el espacio grupal o esconderse bajo los muebles"
                    " cuando se requiere hablar en público."
                ),
                "ex": "Hiding under the back table during science presentation.",
                "ex_zh": "在科学课展示期间躲在后排桌子底下。",
                "ex_es": (
                    "Esconderse bajo la mesa trasera durante la presentación de"
                    " ciencias."
                ),
                "non_ex": "Sitting quietly while listening to a peer present.",
                "non_ex_zh": "安静坐着聆听同伴展示。",
                "non_ex_es": (
                    "Sentarse tranquilamente mientras escucha a un compañero."
                ),
                "dimensions": (
                    "Frequency: 2-3 times per week (4-week trend). Duration:"
                    " 3-8 minutes. Intensity: Moderate."
                ),
                "dimensions_zh": (
                    "频率：每周 2-3 次（4周趋势）。持续时间：3-8 分钟。强度：中度。"
                ),
                "dimensions_es": (
                    "Frecuencia: 2-3 veces por semana (4 semanas). Duración:"
                    " 3-8 min. Intensidad: Moderada."
                ),
                "triggers": (
                    "Setting Events: Social anxiety, public speaking demands."
                ),
                "triggers_zh": "背景事件：社交焦虑、公开演讲要求。",
                "triggers_es": (
                    "Eventos de contexto: Ansiedad social, hablar en público."
                ),
                "consequences": (
                    "Teacher provides alternative individual task option."
                ),
                "consequences_zh": "教师提供个人独立任务替代选项。",
                "consequences_es": (
                    "El maestro ofrece opción de tarea individual alternativa."
                ),
                "qabf_summary": (
                    "Task Escape: 13/15 | Anxiety/Avoidance: 12/15 | Attention:"
                    " 2/15"
                ),
                "qabf_summary_zh": "逃避任务: 13/15 | 焦虑/回避: 12/15 | 社交关注: 2/15",
                "qabf_summary_es": (
                    "Escape: 13/15 | Ansiedad/Evitación: 12/15 | Atención: 2/15"
                ),
                "triangulation": (
                    "65% ABC Data (4 weeks) + 25% Teacher Interview + 10%"
                    " QABF."
                ),
                "triangulation_zh": (
                    "65% ABC 数据（4周） + 25% 教师访谈 + 10% QABF。"
                ),
                "triangulation_es": (
                    "65% Datos ABC (4 semanas) + 25% Entrevista + 10% QABF."
                ),
                "hypothesis": (
                    "Primary Function: Escape from Social Evaluative Stress."
                ),
                "hypothesis_zh": "核心行为功能：逃避社交评估压力。",
                "hypothesis_es": (
                    "Función principal: Escape de estrés evaluativo social."
                ),
                "ferb": (
                    "Request pre-recorded video presentation submission instead"
                    " of live speaking."
                ),
                "ferb_zh": "请求提交预先录制的视频展示替代现场发言。",
                "ferb_es": (
                    "Solicitar presentación en video pregrabado en lugar de en"
                    " vivo."
                ),
            },
            {
                "name": "3. Auditory Overload & Elopement to Hallway",
                "name_zh": "3. 听觉超载与冲向走廊行为",
                "name_es": "3. Sobrecarga auditiva y fuga al pasillo",
                "def": (
                    "Covering ears and running out of the classroom into the"
                    " hallway upon exposure to high ambient noise (>75 dB)."
                    " Onset: hands to ears and bolting; Offset: calming in"
                    " quiet area."
                ),
                "def_zh": (
                    "暴露于高环境噪音（>75 dB）时，双手捂耳并跑出教室进入走廊。以双手捂耳并冲出门为行为开始，以在安静区域平静为行为结束。"
                ),
                "def_es": (
                    "Cubrirse los orejas y salir corriendo del salón al pasillo"
                    " ante ruido ambiental alto."
                ),
                "ex": "Running out to hallway during noisy cafeteria transition.",
                "ex_zh": "在嘈杂的餐厅转换期间跑向走廊。",
                "ex_es": (
                    "Correr al pasillo durante la transición ruidosa de la"
                    " cafetería."
                ),
                "non_ex": "Putting on headphones proactively in quiet room.",
                "non_ex_zh": "在安静房间主动戴上耳机。",
                "non_ex_es": (
                    "Ponerse auriculares proactivamente en una habitación"
                    " silenciosa."
                ),
                "dimensions": (
                    "Frequency: 1-3 times weekly (4-week window). Duration:"
                    " 5-10 minutes. Intensity: Moderate."
                ),
                "dimensions_zh": (
                    "频率：每周 1-3 次（4周窗口期）。持续时间：5-10 分钟。强度：中度。"
                ),
                "dimensions_es": (
                    "Frecuencia: 1-3 veces por semana (4 semanas). Duración:"
                    " 5-10 min. Intensidad: Moderada."
                ),
                "triggers": "Setting Events: Cafeteria, assembly halls, gym classes.",
                "triggers_zh": "背景事件：食堂、礼堂、体育课。",
                "triggers_es": "Eventos de contexto: Cafetería, asambleas, gimnasio.",
                "consequences": (
                    "Aide escorts student to quiet library corner with pass."
                ),
                "consequences_zh": "助教持通行证护送学生至图书馆安静角落。",
                "consequences_es": (
                    "El asistente acompaña al estudiante a un rincón tranquilo."
                ),
                "qabf_summary": (
                    "Sensory Automatic: 14/15 | Escape: 10/15 | Attention:"
                    " 1/15"
                ),
                "qabf_summary_zh": (
                    "感官自动强化: 14/15 | 逃避任务: 10/15 | 社交关注: 1/15"
                ),
                "qabf_summary_es": (
                    "Sensorial automático: 14/15 | Escape: 10/15 | Atención:"
                    " 1/15"
                ),
                "triangulation": (
                    "65% ABC Data (4 weeks) + 25% Aide Interview + 10% QABF."
                ),
                "triangulation_zh": (
                    "65% ABC 数据（4周） + 25% 助教访谈 + 10% QABF。"
                ),
                "triangulation_es": (
                    "65% Datos ABC (4 semanas) + 25% Entrevista + 10% QABF."
                ),
                "hypothesis": (
                    "Primary Function: Automatic Sensory Regulation / Escape"
                    " from Noise."
                ),
                "hypothesis_zh": "核心行为功能：自动感官调节/逃避噪音。",
                "hypothesis_es": (
                    "Función principal: Regulación sensorial / Escape del"
                    " ruido."
                ),
                "ferb": (
                    "Independently retrieve and wear noise-canceling headphones"
                    " upon entering loud areas."
                ),
                "ferb_zh": "进入嘈杂区域时独立取出并佩戴降噪耳机。",
                "ferb_es": (
                    "Tomar y usar auriculares con cancelación de ruido de"
                    " forma independiente."
                ),
            },
        ],
        "strengths": (
            "Excellent visual-spatial abilities, enthusiastic about technology"
            " and drawing."
        ),
        "strengths_zh": "具备出色的视觉空间能力，对科技和绘画抱有极高热情。",
        "strengths_es": (
            "Excelentes habilidades visoespaciales, entusiasta de la"
            " tecnología y el dibujo."
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
                " community living across 4 weeks."
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
                "ex_zh": "当提高工作配额时，大叫“绝不可能！”并将组装盒重重摔在桌上。",
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
                    "Frequency: 1-2 times weekly (tracked over 4 weeks)."
                    " Duration: 5-10 minutes. Intensity: Moderate."
                ),
                "dimensions_zh": (
                    "频率：每周 1-2 次（4周追踪）。持续时间：5-10 分钟。强度：中度。"
                ),
                "dimensions_es": (
                    "Frecuencia: 1-2 veces por semana (4 semanas). Duración:"
                    " 5-10 min. Intensidad: Moderada."
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
                "consequences": "DSP offers choice board, demand temporarily paused.",
                "consequences_zh": (
                    "直属支持人员（DSP）提供选择板，任务要求被暂时暂停。"
                ),
                "consequences_es": (
                    "El personal DSP ofrece panel de opciones, demanda pausada."
                ),
                "qabf_summary": (
                    "Task Escape: 13/15 | Attention: 4/15 | Tangible: 3/15 |"
                    " Sensory: 1/15 | Physical: 1/15"
                ),
                "qabf_summary_zh": (
                    "逃避任务: 13/15 | 社交关注: 4/15 | 获得物质: 3/15 | 感官刺激:"
                    " 1/15 | 身体不适: 1/15"
                ),
                "qabf_summary_es": (
                    "Escape: 13/15 | Atención: 4/15 | Tangible: 3/15 |"
                    " Sensorial: 1/15 | Físico: 1/15"
                ),
                "triangulation": (
                    "65% Direct ABC Data (4 weeks) + 25% Indirect Vocational"
                    " Interview + 10% QABF (Escape score 13/15)."
                ),
                "triangulation_zh": (
                    "65% 直接 ABC 数据（4周） + 25% 间接职业能力访谈 + 10% QABF 结果（逃避得分 13/15）。"
                ),
                "triangulation_es": (
                    "65% Datos ABC (4 semanas) + 25% Entrevista + 10% QABF"
                    " (Escape 13/15)."
                ),
                "hypothesis": (
                    "Primary Function: Escape from Vocational Assembly"
                    " Demands."
                ),
                "hypothesis_zh": "核心行为功能：逃避职业组装工作要求。",
                "hypothesis_es": "Función principal: Escape de demandas vocacionales.",
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
            },
            {
                "name": "2. Property Disruption During Peer Conflict",
                "name_zh": "2. 同伴冲突期间的物品破坏行为",
                "name_es": "2. Alteración de la propiedad durante conflicto con pares",
                "def": (
                    "Throwing objects or sweeping items off surfaces when"
                    " disputes arise over shared community spaces or items."
                    " Onset: grabbing/throwing object; Offset: 2 minutes of calm."
                ),
                "def_zh": (
                    "当在共享社区空间或物品产生争议时，投掷物品或将物品从表面扫落。以抓取/投掷物品为行为开始，以持续 2 分钟平静为行为结束。"
                ),
                "def_es": (
                    "Lojar objetos o barrer elementos de las superficies cuando"
                    " surgen disputas."
                ),
                "ex": "Throwing couch cushions when peer takes remote control.",
                "ex_zh": "当同伴拿走遥控器时扔掉沙发布艺垫。",
                "ex_es": (
                    "Lanzar cojines del sofá cuando un compañero toma el control"
                    " remoto."
                ),
                "non_ex": "Asking staff to mediate peer disagreement calmly.",
                "non_ex_zh": "平静地请求工作人员调解同伴意见分歧。",
                "non_ex_es": (
                    "Pedir al personal que medie en el desacuerdo con calma."
                ),
                "dimensions": (
                    "Frequency: 1-2 times weekly (4-week tracking). Duration:"
                    " 2-5 minutes. Intensity: Moderate."
                ),
                "dimensions_zh": (
                    "频率：每周 1-2 次（4周追踪）。持续时间：2-5 分钟。强度：中度。"
                ),
                "dimensions_es": (
                    "Frecuencia: 1-2 veces por semana (4 semanas). Duración:"
                    " 2-5 min. Intensidad: Moderada."
                ),
                "triggers": (
                    "Setting Events: Shared living spaces, competing for media"
                    " items."
                ),
                "triggers_zh": "背景事件：共享生活空间、争夺娱乐设备。",
                "triggers_es": (
                    "Eventos de contexto: Espacios compartidos, competencia por"
                    " medios."
                ),
                "consequences": (
                    "Staff intervenes, separates individuals, offers"
                    " mediation."
                ),
                "consequences_zh": "工作人员干预，分隔双方并提供调解。",
                "consequences_es": (
                    "El personal interviene, separa a las personas y ofrece"
                    " mediación."
                ),
                "qabf_summary": (
                    "Tangible Access: 14/15 | Escape: 8/15 | Attention: 5/15"
                ),
                "qabf_summary_zh": "获得物质: 14/15 | 逃避任务: 8/15 | 社交关注: 5/15",
                "qabf_summary_es": (
                    "Acceso a tangible: 14/15 | Escape: 8/15 | Atención: 5/15"
                ),
                "triangulation": (
                    "65% ABC Data (4 weeks) + 25% DSP Interview + 10% QABF."
                ),
                "triangulation_zh": (
                    "65% ABC 数据（4周） + 25% DSP 访谈 + 10% QABF。"
                ),
                "triangulation_es": (
                    "65% Datos ABC (4 semanas) + 25% Entrevista + 10% QABF."
                ),
                "hypothesis": (
                    "Primary Function: Access to Tangible / Control over Shared"
                    " Media."
                ),
                "hypothesis_zh": "核心行为功能：获取物质/控制共享媒介。",
                "hypothesis_es": (
                    "Función principal: Acceso a tangible / Control de medios."
                ),
                "ferb": (
                    "Use communication board to negotiate turn-taking schedule"
                    " with peers."
                ),
                "ferb_zh": "使用沟通板与同伴协商轮流使用时间表。",
                "ferb_es": (
                    "Usar tablero de comunicación para negociar turnos con"
                    " pares."
                ),
            },
            {
                "name": "3. Elopement from Facility During Schedule Shifts",
                "name_zh": "3. 日程调整期间擅离社区机构",
                "name_es": "3. Fuga de la instalación durante cambios de horario",
                "def": (
                    "Leaving designated day program perimeter without staff"
                    " notification during unplanned alarms or schedule"
                    " transitions. Onset: crossing perimeter exit; Offset:"
                    " staff intervention."
                ),
                "def_zh": (
                    "在未预料的警报或日程转换期间，未经工作人员通知擅自离开指定的日间项目边界。以跨越边界出口为行为开始，以工作人员干预为行为结束。"
                ),
                "def_es": (
                    "Abandonar el perímetro designado del programa diurno sin"
                    " notificar al personal."
                ),
                "ex": "Running out door during unexpected fire drill alarm.",
                "ex_zh": "在突发消防演习警报期间跑出门外。",
                "ex_es": (
                    "Correr hacia afuera durante una alarma de simulacro de"
                    " incendio."
                ),
                "non_ex": "Walking in orderly line with group during scheduled drills.",
                "non_ex_zh": "在计划演习期间随队伍有序排队行走。",
                "non_ex_es": (
                    "Caminar en fila ordenada con el grupo durante simulacros"
                    " programados."
                ),
                "dimensions": (
                    "Frequency: 1 time monthly (monitored across 4-week blocks)."
                    " Duration: 10-15 minutes. Intensity: Moderate to High."
                ),
                "dimensions_zh": (
                    "频率：每月 1 次（4周观察周期监测）。持续时间：10-15 分钟。强度：中度至高度。"
                ),
                "dimensions_es": (
                    "Frecuencia: 1 vez al mes (bloques de 4 semanas). Duración:"
                    " 10-15 min. Intensidad: Moderada a Alta."
                ),
                "triggers": (
                    "Setting Events: Sudden loud alarms, unexpected transition"
                    " triggers."
                ),
                "triggers_zh": "背景事件：突发高分贝警报、意外转换触发因素。",
                "triggers_es": (
                    "Eventos de contexto: Alarmas fuertes repentinas,"
                    " transiciones inesperadas."
                ),
                "consequences": (
                    "Staff escorts client safely back and reviews visual"
                    " schedule."
                ),
                "consequences_zh": "工作人员护送客户安全返回并复习视觉日程表。",
                "consequences_es": (
                    "El personal acompaña al cliente de regreso y revisa el"
                    " horario visual."
                ),
                "qabf_summary": (
                    "Escape: 13/15 | Sensory/Anxiety: 11/15 | Attention: 2/15"
                ),
                "qabf_summary_zh": "逃避任务: 13/15 | 感官/焦虑: 11/15 | 社交关注: 2/15",
                "qabf_summary_zh_es": (
                    "Escape: 13/15 | Sensorial/Ansiedad: 11/15 | Atención:"
                    " 2/15"
                ),
                "triangulation": (
                    "65% ABC Data (4 weeks) + 25% Staff Interview + 10% QABF."
                ),
                "triangulation_zh": (
                    "65% ABC 数据（4周） + 25% 工作人员访谈 + 10% QABF。"
                ),
                "triangulation_es": (
                    "65% Datos ABC (4 semanas) + 25% Entrevista + 10% QABF."
                ),
                "hypothesis": (
                    "Primary Function: Escape from Sudden Sensory/Alarm"
                    " Overload."
                ),
                "hypothesis_zh": "核心行为功能：逃避突发感官/警报超载。",
                "hypothesis_es": (
                    "Función principal: Escape de sobrecarga"
                    " sensorial/alarma repentina."
                ),
                "ferb": (
                    "Request staff accompaniment or quiet transition space using"
                    " card."
                ),
                "ferb_zh": "使用卡片请求工作人员陪同或安静的转换空间。",
                "ferb_es": (
                    "Solicitar acompañamiento del personal o espacio"
                    " tranquilo con tarjeta."
                ),
            },
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
            "Participates in servicios vocacionales para adultos bajo exención"
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

# ==========================================
# PHASE EXCLUSIVITY & BLUEPRINT WARNING (Requirement 2)
# ==========================================
st.warning(
    "🚨 **Phase 1 (Demo Version) vs. Phase 2 Architecture Notice**\n\n"
    "• **Current Status (Phase 1 Demo):** This browser session runs exclusively on **Mock Data**. No patient or client health information is stored or uploaded to any external server.\n\n"
    "• **Phase 2 Local Deployment Blueprint:** In production deployment, this tool will execute **locally on the BCBA's secure offline computer**. "
    "All uploaded client notes will undergo **automatic on-device de-identification via local regex scripts** (replacing identifiers with placeholders like `[CLIENT_NAME]` and `[DOB]`). "
    "For instance, a BCBA can export generated drafts and perform a quick local **Ctrl + H** find-and-replace sweep prior to clinical sign-off.",
    icon="⚠️",
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

st.markdown(
    f"""
    <div class="protocol-card">
        <div class="protocol-title">📋 Selected Protocol Framework: {current_meta['title']} ({current_meta['framework']})</div>
        {"".join([f'<div class="protocol-bullet">{s}</div>' for s in current_meta['protocol_sentences']])}
    </div>
""",
    unsafe_allow_html=True,
)

col_input1, col_input2, col_input3, col_input4 = st.columns([1, 1, 1, 1])

with col_input1:
  st.markdown("#### 📄 Direct Observation (ABC)")
  mock_csv = generate_mock_abc_csv(selected_cohort_key)
  st.download_button(
      label=f"📥 Download Mock ABC (.csv)",
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
      label=f"📥 Download Mock Interview (.docx)",
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
  st.markdown("#### 📊 Behavior QABF Assessment")
  mock_qabf = generate_mock_qabf_docx(selected_cohort_key)
  st.download_button(
      label=f"📥 Download Mock QABF (.docx)",
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

with col_input4:
  st.markdown("#### 📈 Behavior Tracking Doc")
  mock_tracking = generate_mock_tracking_csv(selected_cohort_key)
  st.download_button(
      label=f"📥 Download Mock Tracking (.csv)",
      data=mock_tracking,
      file_name=f"DeIdentified_Tracking_{current_meta['file_tag']}.csv",
      mime="text/csv",
      use_container_width=True,
  )
  uploaded_tracking = st.file_uploader(
      "Upload Tracking File:",
      type=["csv", "xlsx"],
      key=f"tracking_{selected_cohort_key}",
  )

active_behaviors = current_meta["behaviors"]


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
        (
            "Client ID",
            "[CLIENT_ID]",
            "Assessment Date",
            "2026-08-14",
        ),
        (
            "Facility/School",
            "[DISTRICT_OR_FACILITY_NAME]",
            "Setting",
            c_meta["setting_str"],
        ),
        (
            "Assessor",
            "[BCBA_NAME], BCBA, LBA",
            "Framework",
            c_meta["framework"],
        ),
        (
            "Primary Language",
            "English",
            "Informants",
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

  lang_mode = "en"
  if "Chinese" in lang_choice:
    lang_mode = "zh"
  elif "Spanish" in lang_choice:
    lang_mode = "es"

  p_t = doc.add_paragraph()
  r_t = p_t.add_run("FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT")
  r_t.bold = True
  p_t.style = doc.styles["Title"]

  if lang_mode == "zh":
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[功能性行为评估 (FBA) 报告 Draft]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)
  elif lang_mode == "es":
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[Informe de Evaluación de la Conducta Funcional (FBA)]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)

  # Section 1: Demographics
  add_bi_heading(
      doc,
      1,
      "1. Student Demographics & Administrative Info",
      (
          "1. 学生/客户基本信息与行政登记"
          if lang_mode == "zh"
          else "1. Datos demográficos del estudiante"
          if lang_mode == "es"
          else None
      ),
  )
  build_compact_demographics_table(doc, c_meta, lang_mode)

  # Section 2: Data Sources & Triangulation Methodology
  add_bi_heading(
      doc,
      1,
      "2. Data Sources & Triangulation Methodology",
      (
          "2. 数据来源与三方交叉验证评估方法"
          if lang_mode == "zh"
          else "2. Fuentes de datos y metodología de triangulación"
          if lang_mode == "es"
          else None
      ),
  )
  add_bi_item(
      doc,
      "Methodology & Triangulation Formula",
      (
          "Triangulation Algorithm Standard Applied: 65% Direct ABC Data + 25%"
          " Indirect Interview Data + 10% QABF (supplemented by Longitudinal"
          " 4-Week Behavior Tracking Document).\n\n1. Direct ABC Data: Continuous"
          " recording across 4-week baseline therapy sessions.\n2. Indirect Assessment:"
          " Structured interviews with parent and lead RBT.\n3. Psychometric"
          " Rating Scale: Behavior-specific QABF (Questions About Behavioral"
          " Function).\n4. Behavior Tracking Document: 4-week longitudinal frequency"
          " and duration monitoring."
      ),
      (
          "评估方法与三方验证公式"
          if lang_mode == "zh"
          else "Metodología y fórmula de triangulación"
          if lang_mode == "es"
          else None
      ),
      (
          "应用四星期三方数据加权验证算法：65% Direct ABC Data + 25% Indirect Interview Data"
          " + 10% QABF（辅以 4 周行为追踪长期记录文件）。\n\n1. 直接 ABC 数据：4 周干预课期间连续行为观察记录。\n2."
          " 间接评估：与家长和督导 RBT 的结构化访谈。\n3. 心理测量量表：针对具体行为的 QABF"
          " 评估问卷。\n4. 行为追踪文件：4 周纵向频率与持续时间监测。"
          if lang_mode == "zh"
          else "Algoritmo de triangulación (4 semanas): 65% ABC + 25% Entrevista + 10%"
          " QABF (con documento de seguimiento longitudinal de 4 semanas)."
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  # Section 3: Background & Strengths
  add_bi_heading(
      doc,
      1,
      "3. Brief Background & Strengths Summary",
      (
          "3. 学生背景与优势摘要"
          if lang_mode == "zh"
          else "3. Antecedentes y resumen de fortalezas"
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
          else "Fortalezas y preferencias"
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
  add_bi_item(
      doc,
      "Clinical / Educational History",
      c_meta["history"],
      (
          "临床/教育背景"
          if lang_mode == "zh"
          else "Historial clínico / educativo"
          if lang_mode == "es"
          else None
      ),
      (
          c_meta["history_zh"]
          if lang_mode == "zh"
          else c_meta.get("history_es", c_meta["history"])
      ),
      lang_mode,
  )

  # Section 4: Individual Functional Analyses (Behavior-by-Behavior with Dedicated Chart)
  add_bi_heading(
      doc,
      1,
      "4. Individual Target Behavior Functional Analyses",
      (
          "4. 目标行为独立功能分析 (按行为逐项拆解与专属配图)"
          if lang_mode == "zh"
          else "4. Análisis funcional individual por conducta y gráficos exclusivos"
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

    # Insert Dedicated Behavior-Specific Tracking Chart (4-week)
    chart_title_en = f"Figure 4.{idx}. 4-Week Longitudinal Behavior Tracking Trend for Target Behavior #{idx}"
    chart_title_zh = f"图 4.{idx}. 目标行为 #{idx} 4周专属纵向追踪趋势图"
    chart_title_es = f"Figura 4.{idx}. Tendencia de seguimiento de 4 semanas para Conducta #{idx}"

    add_bi_heading(
        doc,
        3,
        chart_title_en,
        chart_title_zh
        if lang_mode == "zh"
        else chart_title_es
        if lang_mode == "es"
        else None,
    )
    chart_buf = generate_behavior_tracking_chart(cohort_key, idx - 1)
    doc.add_picture(chart_buf, width=Inches(5.0))
    doc.add_paragraph(
        f"Note: The chart above illustrates the 4-week independent tracking frequency"
        f" trend specifically recorded for Target Behavior #{idx}."
    )

    add_bi_item(
        doc,
        "A. Operational Definition",
        b["def"],
        (
            "A. 操作性定义"
            if lang_mode == "zh"
            else "A. Definición operacional"
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
        "B. Examples & Non-Examples",
        f"Examples: {b['ex']}\nNon-Examples: {b['non_ex']}",
        (
            "B. 示例与非示例"
            if lang_mode == "zh"
            else "B. Ejemplos y no ejemplos"
            if lang_mode == "es"
            else None
        ),
        (
            f"示例: {b.get('ex_zh')}\n非示例: {b.get('non_ex_zh')}"
            if lang_mode == "zh"
            else f"Ejemplos: {b.get('ex_es')}\nNo ejemplos: {b.get('non_ex_es')}"
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )
    add_bi_item(
        doc,
        "C. Behavior Dimensions",
        b["dimensions"],
        (
            "C. 行为维度 (频率/持续时间/强度)"
            if lang_mode == "zh"
            else "C. Dimensiones de la conducta"
            if lang_mode == "es"
            else None
        ),
        (
            b.get("dimensions_zh")
            if lang_mode == "zh"
            else b.get("dimensions_es")
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )
    add_bi_item(
        doc,
        "D. Environmental Triggers & Setting Events",
        b["triggers"],
        (
            "D. 环境触发因素与背景事件"
            if lang_mode == "zh"
            else "D. Desencadenantes y eventos de contexto"
            if lang_mode == "es"
            else None
        ),
        (
            b.get("triggers_zh")
            if lang_mode == "zh"
            else b.get("triggers_es")
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )
    add_bi_item(
        doc,
        "E. Maintaining Consequences",
        b["consequences"],
        (
            "E. 维持后果与他人反应"
            if lang_mode == "zh"
            else "E. Consecuencias de mantenimiento"
            if lang_mode == "es"
            else None
        ),
        (
            b.get("consequences_zh")
            if lang_mode == "zh"
            else b.get("consequences_es")
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )
    add_bi_item(
        doc,
        "F. Behavior-Specific QABF Results",
        b["qabf_summary"],
        (
            "F. 该行为专属 QABF 量表得分"
            if lang_mode == "zh"
            else "F. Resultados QABF específicos"
            if lang_mode == "es"
            else None
        ),
        (
            b.get("qabf_summary_zh")
            if lang_mode == "zh"
            else b.get("qabf_summary_es")
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )
    add_bi_item(
        doc,
        "G. Triangulation (65% ABC + 25% Interview + 10% QABF)",
        b["triangulation"],
        (
            "G. 三方交叉验证算法结果"
            if lang_mode == "zh"
            else "G. Triangulación de datos"
            if lang_mode == "es"
            else None
        ),
        (
            b.get("triangulation_zh")
            if lang_mode == "zh"
            else b.get("triangulation_es")
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )
    add_bi_item(
        doc,
        "H. Hypothesized Function",
        b["hypothesis"],
        (
            "H. 该行为推断功能"
            if lang_mode == "zh"
            else "H. Función hipetotizada"
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

  # Section 5: Overall Synthesis
  add_bi_heading(
      doc,
      1,
      "5. Synthesis & Personalised Clinical Recommendations",
      (
          "5. 个性化综合评估结论与临床建议"
          if lang_mode == "zh"
          else "5. Síntesis y recomendaciones clínicas personalizadas"
          if lang_mode == "es"
          else None
      ),
  )

  if cohort_key == "g1":
    synth_en = (
        "Based on the 4-week triangulated functional assessment (65% Direct ABC + 25%"
        " Indirect Interview + 10% QABF), the client (Ages 2-5) primarily"
        " demonstrates behaviors maintained by Social Negative Reinforcement"
        " (Task/Transition Escape) and Sensory Automatic Regulation"
        " (Auditory Aversion). Recommendations:\n1. Formulate an NDBI / ESDM"
        " play-based Behavior Intervention Plan (BIP) integrating AAC (BigMack"
        " / PECS 'Break' card) for early functional communication training"
        " (FCT).\n2. Implement antecedent noise-mitigation protocols"
        " (noise-canceling headphones, visual schedule warnings) prior to"
        " structured table tasks and transitions.\n3. Train parents and RBTs on"
        " shared co-regulation strategies to replace self-injurious and"
        " sensory distress behaviors."
    )
    synth_zh = (
        "基于 4 周 65% 直接 ABC 观察 + 25% 间接访谈 + 10% QABF 量表的三方交叉验证算法，该 2-5"
        " 岁幼儿的目标行为主要由‘逃避任务/转换（社交负强化）’及‘调节听觉过度刺激（感官自动强化）’所维持。临床建议：\n1."
        " 制定基于 ESDM / NDBI 自然教法的行为干预计划 (BIP)，优先引入 AAC"
        " 语音按键与 PECS‘休息’视觉卡进行早期功能性沟通训练 (FCT)。\n2."
        " 在结构化桌面任务与环节转换前，实施前因降噪与视觉倒计时预警。\n3."
        " 对家长及 RBT 进行情绪共调节 (Co-regulation) 技巧培训，以替代自伤与感官情绪发作。"
    )
    synth_es = (
        "Con base en la triangulación de 4 semanas (65% ABC + 25% Entrevista +"
        " 10% QABF), el cliente (2-5 años) presenta conductas mantenidas por"
        " Escape de Tareas y Regulación Sensorial Auditiva. Recomendaciones:\n1."
        " Diseñar un BIP basado en NDBI/ESDM integrando AAC/PECS para Comunicación"
        " Funcional (FCT).\n2. Implementar protocolos de mitigación de ruido y"
        " avisos visuales antes de transiciones."
    )
  elif cohort_key == "g2":
    synth_en = (
        "Based on the 4-week triangulated functional assessment, the student (Ages 5-21)"
        " demonstrates academic escape and sensory avoidance behaviors. Recommendations:\n1."
        " Formulate an IEP-aligned PBIS Behavior Intervention Plan (BIP) incorporating"
        " structured breaks and self-advocacy cards.\n2. Provide proactive visual schedules"
        " and pre-recorded presentation options to reduce anxiety."
    )
    synth_zh = (
        "基于 4 周三方交叉验证评估，该学龄学生表现出学业逃避与感官回避行为。临床建议：\n1. 制定符合 IEP 标准的 PBIS"
        " 行为干预计划 (BIP)，纳入结构化休息与自我倡导卡片。\n2. 提供前置视觉日程表及预录制展示选项以降低焦虑。"
    )
    synth_es = (
        "Con base en la evaluación de 4 semanas, el estudiante presenta"
        " evitación académica y sensorial. Recomendaciones:\n1. Diseñar un BIP"
        " alineado con IEP/PBIS.\n2. Proveer horarios visuales y opciones"
        " alternativas."
    )
  else:
    synth_en = (
        "Based on the 4-week triangulated functional assessment, the adult client (21+) exhibits"
        " vocational task refusal and property disruption maintained by escape and tangible access."
        " Recommendations:\n1. Implement person-centered adult day program strategies."
        " \n2. Provide choice boards and self-advocacy training for collaborative scheduling."
    )
    synth_zh = (
        "基于 4 周三方交叉验证评估，该成年客户表现出由逃避和获得物质维持的职业任务拒绝及物品破坏行为。临床建议：\n1."
        " 落实以人为本的成人日间项目策略。\n2. 提供选择板及自我倡导培训以协助协商日常时间表。"
    )
    synth_es = (
        "Con base en la evaluación de 4 semanas, el cliente adulto exhibe"
        " rechazo vocacional. Recomendaciones:\n1. Implementar estrategias"
        " centradas en la persona.\n2. Proveer tableros de elección."
    )

  add_bi_item(
      doc,
      "Clinical Synthesis & Recommendations",
      synth_en,
      (
          "临床综合结论与建议"
          if lang_mode == "zh"
          else "Síntesis clínica y recomendaciones"
          if lang_mode == "es"
          else None
      ),
      synth_zh if lang_mode == "zh" else synth_es if lang_mode == "es" else None,
      lang_mode,
  )

  # ==========================================
  # CLINICAL RESPONSIBILITY NOTICE (Requirement 3)
  # ==========================================
  add_bi_heading(
      doc,
      1,
      "6. Clinical Responsibility & Tool Disclaimer Notice",
      (
          "6. 临床责任与工具免责声明"
          if lang_mode == "zh"
          else "6. Responsabilidad clínica y aviso legal"
          if lang_mode == "es"
          else None
      ),
  )
  notice_text_en = (
      "CLINICAL RESPONSIBILITY NOTICE:\n"
      "This document and its associated outputs function exclusively as a FIRST-DRATCH SYNTHESIZER "
      "designed to assist Board Certified Behavior Analysts (BCBAs) and Licensed Behavior Analysts (LBAs) "
      "in drafting functional behavior assessments and behavioral intervention plans.\n"
      "FINAL SIGN-OFF, PROFESSIONAL VALIDATION, AND ENTIRE CLINICAL RESPONSIBILITY (CLINICAL RESPONSIBILITY) "
      "REST SOLELY AND EXCLUSIVELY WITH THE QUALIFIED BCBA OR LBA SIGNING AND IMPLEMENTING THIS PLAN. "
      "This tool does not replace professional clinical judgment or direct patient evaluation."
  )
  notice_text_zh = (
      "临床责任声明：\n"
      "本文件及相关生成结果仅作为“初稿合成器 (First-Draft Synthesizer)”使用，旨在协助认证行为分析师 (BCBA) "
      "与持证行为分析师 (LBA) 起草功能性行为评估与行为干预计划。\n"
      "最终签字确认、专业效度验证以及所有临床责任 (Clinical Responsibility) "
      "均完全且专属於签署并实施该计划的合格 BCBA 或 LBA。本工具不能替代专业的临床判断或直接的患者评估。"
  )
  notice_text_es = (
      "AVISO DE RESPONSABILIDAD CLÍNICA:\n"
      "Este documento funciona exclusivamente como un SINTETIZADOR DE BORRADOR INICIAL (First-Draft Synthesizer). "
      "LA RESPONSABILIDAD CLÍNICA FINAL (Clinical Responsibility) recae única y exclusivamente en el BCBA o LBA calificado."
  )
  add_bi_item(
      doc,
      "Tool Positioning & Liability",
      notice_text_en,
      "工具定位与责任归属" if lang_mode == "zh" else "Posicionamiento y responsabilidad" if lang_mode == "es" else None,
      notice_text_zh if lang_mode == "zh" else notice_text_es if lang_mode == "es" else None,
      lang_mode,
  )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


def generate_exact_bip_doc(cohort_key, lang_choice, behavior_list):
  c_meta = cohort_meta[cohort_key]
  doc = docx.Document()

  lang_mode = "en"
  if "Chinese" in lang_choice:
    lang_mode = "zh"
  elif "Spanish" in lang_choice:
    lang_mode = "es"

  p_t = doc.add_paragraph()
  r_t = p_t.add_run("BEHAVIOR INTERVENTION PLAN (BIP)")
  r_t.bold = True
  p_t.style = doc.styles["Title"]

  if lang_mode == "zh":
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[行为干预计划 (BIP) 报告 Draft]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)
  elif lang_mode == "es":
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run("[Plan de Intervención Conductual (BIP)]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)

  # Section 1: Demographics
  add_bi_heading(
      doc,
      1,
      "1. Student Demographics & Plan Metadata",
      (
          "1. 学生基本信息与计划元数据"
          if lang_mode == "zh"
          else "1. Datos demográficos y metadatos"
          if lang_mode == "es"
          else None
      ),
  )
  build_compact_demographics_table(doc, c_meta, lang_mode)

  # Section 2: Proactive / Antecedent Strategies (4-week framework)
  add_bi_heading(
      doc,
      1,
      "2. Proactive & Antecedent Modification Strategies (4-Week Protocol)",
      (
          "2. 前因修改与预防性策略 (4周观察周期标准)"
          if lang_mode == "zh"
          else "2. Estrategias proactivas y antecedentes (4 semanas)"
          if lang_mode == "es"
          else None
      ),
  )
  add_bi_item(
      doc,
      "Antecedent Interventions",
      (
          "1. Visual Schedules & Priming: Provide 5-minute visual warnings prior to activity transitions across 4-week structured intervals.\n"
          "2. Environmental Noise Mitigation: Access to noise-canceling headphones during high ambient auditory stimulation.\n"
          "3. Choice-Making Opportunities: Offer structured choices regarding academic or vocational task order."
      ),
      (
          "前因干预策略"
          if lang_mode == "zh"
          else "Intervenciones antecedentes"
          if lang_mode == "es"
          else None
      ),
      (
          "1. 视觉日程表与预告：在 4 周结构化干预期间，于活动转换前 5 分钟提供视觉警示。\n"
          "2. 环境降噪：在高分贝声音刺激环境中提供降噪耳机。\n"
          "3. 自主选择机会：就学业或职业任务顺序提供结构化选择。"
          if lang_mode == "zh"
          else "1. Horarios visuales. 2. Mitigación de ruido. 3. Opciones."
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  # Section 3: Replacement Behaviors (FERB)
  add_bi_heading(
      doc,
      1,
      "3. Functional Replacement Behaviors (FERB)",
      (
          "3. 功能性替代行为 (FERB) 训练计划"
          if lang_mode == "zh"
          else "3. Conductas de reemplazo funcional (FERB)"
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
        f"Target Behavior #{idx}: {b['name']} - FERB",
        f"目标行为 #{idx}: {trans_name} - FERB" if trans_name else None,
    )
    add_bi_item(
        doc,
        "Replacement Protocol",
        b["ferb"],
        (
            "替代行为方案"
            if lang_mode == "zh"
            else "Protocolo de reemplazo"
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

  # Section 4: Consequence & Reinforcement Strategies
  add_bi_heading(
      doc,
      1,
      "4. Consequence & Reinforcement Protocols",
      (
          "4. 后果处理与强化协议"
          if lang_mode == "zh"
          else "4. Protocolos de consecuencia y refuerzo"
          if lang_mode == "es"
          else None
      ),
  )
  add_bi_item(
      doc,
      "Differential Reinforcement & Error Correction",
      (
          "1. Extinction / Planned Ignoring: Do not reinforce problem behaviors with task escape or unauthorized tangibles.\n"
          "2. Differential Reinforcement of Alternative Behavior (DRA): Immediately reinforce functional communication (AAC/PECS/cards).\n"
          "3. Safety First: Always prioritize physical safety, blocking, and non-punitive crisis de-escalation."
      ),
      (
          "区别性强化与纠错"
          if lang_mode == "zh"
          else "Refuerzo diferencial y corrección"
          if lang_mode == "es"
          else None
      ),
      (
          "1. 消退/计划性忽略：切勿通过逃避任务或未经授权的物品来强化问题行为。\n"
          "2. 替代行为区别性强化 (DRA)：对有效沟通（AAC/PECS/卡片）给予即时强化。\n"
          "3. 安全第一：始终将人身安全、物理阻挡及非惩罚性危机降级放在首位。"
          if lang_mode == "zh"
          else "1. Extinción. 2. DRA. 3. Seguridad."
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  # Section 5: Clinical Responsibility Notice (Requirement 3)
  add_bi_heading(
      doc,
      1,
      "5. Clinical Responsibility & Tool Disclaimer Notice",
      (
          "5. 临床责任与工具免责声明"
          if lang_mode == "zh"
          else "5. Responsabilidad clínica y aviso legal"
          if lang_mode == "es"
          else None
      ),
  )
  notice_text_en = (
      "CLINICAL RESPONSIBILITY NOTICE:\n"
      "This document and its associated outputs function exclusively as a FIRST-DRAATCH SYNTHESIZER "
      "designed to assist Board Certified Behavior Analysts (BCBAs) and Licensed Behavior Analysts (LBAs) "
      "in drafting functional behavior assessments and behavioral intervention plans.\n"
      "FINAL SIGN-OFF, PROFESSIONAL VALIDATION, AND ENTIRE CLINICAL RESPONSIBILITY (CLINICAL RESPONSIBILITY) "
      "REST SOLELY AND EXCLUSIVELY WITH THE QUALIFIED BCBA OR LBA SIGNING AND IMPLEMENTING THIS PLAN. "
      "This tool does not replace professional clinical judgment or direct patient evaluation."
  )
  notice_text_zh = (
      "临床责任声明：\n"
      "本文件及相关生成结果仅作为“初稿合成器 (First-Draft Synthesizer)”使用，旨在协助认证行为分析师 (BCBA) "
      "与持证行为分析师 (LBA) 起草功能性行为评估与行为干预计划。\n"
      "最终签字确认、专业效度验证以及所有临床责任 (Clinical Responsibility) "
      "均完全且专属於签署并实施该计划的合格 BCBA 或 LBA。本工具不能替代专业的临床判断或直接的患者评估。"
  )
  notice_text_es = (
      "AVISO DE RESPONSABILIDAD CLÍNICA:\n"
      "Este documento funciona exclusivamente como un SINTETIZADOR DE BORRADOR INICIAL (First-Draft Synthesizer). "
      "LA RESPONSABILIDAD CLÍNICA FINAL (Clinical Responsibility) recae única y exclusivamente en el BCBA o LBA calificado."
  )
  add_bi_item(
      doc,
      "Tool Positioning & Liability",
      notice_text_en,
      "工具定位与责任归属" if lang_mode == "zh" else "Posicionamiento y responsabilidad" if lang_mode == "es" else None,
      notice_text_zh if lang_mode == "zh" else notice_text_es if lang_mode == "es" else None,
      lang_mode,
  )

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 8. Interactive Document Generation Section
# ==========================================
st.divider()
st.markdown("### 3️⃣ Generate Clinical Report & BIP First-Drafts")

col_gen1, col_gen2 = st.columns([1, 1])

with col_gen1:
  st.markdown("#### ⚙️ Document Language & Customization")
  lang_choice = st.selectbox(
      "Select Output Report Language / 报告语言:",
      [
          "English (Standard BCBA)",
          "Bilingual (English + 中文 Chinese)",
          "Bilingual (English + Español Spanish)",
      ],
      index=1,
  )

  st.markdown("#### 🎯 Target Behaviors to Include in Synthesis")
  selected_behavior_indices = []
  for idx, b in enumerate(active_behaviors):
    if st.checkbox(f"Include {b['name']}", value=True, key=f"beh_sel_{selected_cohort_key}_{idx}"):
      selected_behavior_indices.append(idx)

  if not selected_behavior_indices:
    st.warning("⚠️ Please select at least one target behavior for report synthesis.")
    filtered_behaviors = active_behaviors
  else:
    filtered_behaviors = [active_behaviors[i] for i in selected_behavior_indices]

with col_gen2:
  st.markdown("#### 📥 Download Compiled Clinical Documents (.docx)")
  st.info(
      "Click below to generate fully formatted, de-identified Microsoft Word (.docx) reports "
      "incorporating 4-week trend tracking charts, separate target behavior breakdowns, and Clinical Responsibility Notices."
  )

  if st.button("🚀 Generate FBA Report Draft", use_container_width=True):
    with st.spinner("Synthesizing 4-week FBA Report Draft..."):
      fba_bio = generate_exact_fba_doc(
          selected_cohort_key, lang_choice, filtered_behaviors
      )
      st.success("✅ FBA Report Draft successfully compiled!")
      st.download_button(
          label="💾 Download FBA Report (.docx)",
          data=fba_bio,
          file_name=f"FBA_Report_{current_meta['file_tag']}.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
          use_container_width=True,
      )

  if st.button("🚀 Generate BIP Document Draft", use_container_width=True):
    with st.spinner("Synthesizing 4-week BIP Document Draft..."):
      bip_bio = generate_exact_bip_doc(
          selected_cohort_key, lang_choice, filtered_behaviors
      )
      st.success("✅ BIP Document Draft successfully compiled!")
      st.download_button(
          label="💾 Download BIP Document (.docx)",
          data=bip_bio,
          file_name=f"BIP_Document_{current_meta['file_tag']}.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
          use_container_width=True,
      )
