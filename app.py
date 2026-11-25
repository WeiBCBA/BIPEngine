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
    page_title="BIP Engine - BCBA FBA & BIP Draft Formulation Tool",
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
              "Date_Time": "2026-08-10 11:00",
              "Setting": "Table Therapy Room",
              "Antecedent": "Presented fine-motor matching worksheet demand",
              "Behavior": (
                  "[TB#1 SIB] Head banging on foam mat (4 forceful contacts)"
                  " and biting right wrist"
              ),
              "Consequence": (
                  "RBT blocked contact with pad/chew, paused demand, prompted"
                  " PECS 'Break'"
              ),
          },
          {
              "Date_Time": "2026-08-11 15:20",
              "Setting": "Clinic Snack Area",
              "Antecedent": "Preferred juice cup removed and table task started",
              "Behavior": (
                  "[TB#1 SIB] Forceful head tap on table edge and biting"
                  " wrist"
              ),
              "Consequence": (
                  "RBT safely blocked contact, paused demand, prompted AAC"
                  " 'I need a break'"
              ),
          },
          {
              "Date_Time": "2026-08-10 09:15",
              "Setting": "Clinic Playroom",
              "Antecedent": "Loud kitchen blender noise started in adjacent room",
              "Behavior": (
                  "[TB#2 Sensory Vocal Distress & Face-Slapping] Screamed"
                  " loudly (>80 dB) and slapped face 3 times"
              ),
              "Consequence": (
                  "RBT offered noise-canceling headphones and redirected to"
                  " sensory chew tool"
              ),
          },
          {
              "Date_Time": "2026-08-12 09:40",
              "Setting": "Sensory Room",
              "Antecedent": "Sudden loud chime sounded from room timer",
              "Behavior": (
                  "[TB#2 Sensory Vocal Distress & Face-Slapping] High-volume"
                  " screaming and striking cheeks"
              ),
              "Consequence": (
                  "RBT provided noise-canceling headphones and guided to"
                  " 'Headphones' symbol"
              ),
          },
          {
              "Date_Time": "2026-08-10 10:30",
              "Setting": "Outdoor Playground",
              "Antecedent": "Transition prompt given to pack up sand toys",
              "Behavior": (
                  "[TB#3 Floor Dropping & Pica Attempts] Dropped to sand and"
                  " attempted to mouth sand"
              ),
              "Consequence": (
                  "RBT blocked hand-to-mouth, offered oral chew, delayed"
                  " transition via '1 More Minute'"
              ),
          },
          {
              "Date_Time": "2026-08-12 11:15",
              "Setting": "Clinic Floor",
              "Antecedent": "Instruction given to clean up blocks for transition",
              "Behavior": (
                  "[TB#3 Floor Dropping & Pica Attempts] Dropped to floor and"
                  " tried to mouth small plastic piece"
              ),
              "Consequence": (
                  "RBT blocked mouthing, provided chew device, prompted '1"
                  " More Minute' card"
              ),
          },
      ],
      "g2": [
          {
              "Date_Time": "2026-08-10 09:30",
              "Setting": "Gen-Ed Classroom",
              "Antecedent": "Teacher denied immediate preferred Chromebook time",
              "Behavior": (
                  "[TB#1 Tantrum and Physical Aggression] Stomped foot, yelled"
                  " loudly, pushed chair, and pushed aide"
              ),
              "Consequence": (
                  "Staff cleared peers, stayed neutral, safely blocked push,"
                  " prompted 'Wait / More Time' when calm"
              ),
          },
          {
              "Date_Time": "2026-08-11 14:30",
              "Setting": "Resource Room",
              "Antecedent": "Asked to stop preferred drawing activity for transition",
              "Behavior": (
                  "[TB#1 Tantrum and Physical Aggression] Slapped desk top"
                  " twice, shouted, and hit arm"
              ),
              "Consequence": (
                  "Staff neutrally blocked hits/slaps, paused demand briefly,"
                  " prompted 'More Time' card"
              ),
          },
          {
              "Date_Time": "2026-08-10 11:00",
              "Setting": "Science Lab",
              "Antecedent": "Presented multi-step lab worksheet without chunking",
              "Behavior": (
                  "[TB#2 Task Refusal and Elopement] Pushed packet away,"
                  " refused task, and exited classroom"
              ),
              "Consequence": (
                  "Staff prompted 'Break/Help' card, safely guided back to"
                  " learning area after brief break"
              ),
          },
          {
              "Date_Time": "2026-08-12 10:00",
              "Setting": "Language Arts",
              "Antecedent": "Assigned 2-page independent writing prompt",
              "Behavior": (
                  "[TB#2 Task Refusal and Elopement] Crossed arms, refused"
                  " pencil, and walked out of learning zone"
              ),
              "Consequence": (
                  "Teacher provided 'Help/Break' card and redirected back to"
                  " small-group table"
              ),
          },
          {
              "Date_Time": "2026-08-11 13:15",
              "Setting": "School Cafeteria",
              "Antecedent": "Sudden loud tray clatter and high cafeteria noise",
              "Behavior": (
                  "[TB#3 Screaming (Auditory Overload)] Screamed at high"
                  " volume with hands near ears"
              ),
              "Consequence": (
                  "Staff offered noise-canceling headphones and guided to"
                  " quiet sensory corner"
              ),
          },
      ],
      "g3": [
          {
              "Date_Time": "2026-08-10 10:15",
              "Setting": "Vocational Workshop",
              "Antecedent": "Supervisor raised box assembly quota by 20 units",
              "Behavior": (
                  "[TB#1 Vocational Task Refusal & Verbal Aggression] Shouted"
                  " 'No way!', refused assembly, slammed parts"
              ),
              "Consequence": (
                  "Job coach offered choice board and 5-minute break, paused"
                  " quota demand"
              ),
          },
          {
              "Date_Time": "2026-08-12 11:00",
              "Setting": "Kitchen Prep Area",
              "Antecedent": "Staff asked client to wash cooking pans immediately",
              "Behavior": (
                  "[TB#1 Vocational Task Refusal & Verbal Aggression] Refused"
                  " task, crossed arms, shouted threats"
              ),
              "Consequence": (
                  "Staff gave space, offered '5-minute break' card, paused"
                  " demand"
              ),
          },
          {
              "Date_Time": "2026-08-11 13:30",
              "Setting": "Community Living Room",
              "Antecedent": "Peer took TV remote control during shared media time",
              "Behavior": (
                  "[TB#2 Property Disruption During Peer Conflict] Grabbed"
                  " couch cushion and threw it across room"
              ),
              "Consequence": (
                  "Staff intervened neutrally, separated dispute, prompted"
                  " turn-taking board"
def generate_mock_interview_docx(cohort_key):
  doc = docx.Document()
  b_list = cohort_meta[cohort_key]["behaviors"]
  tb1 = b_list[0]["name"]
  tb2 = b_list[1]["name"]
  tb3 = b_list[2]["name"]
  doc.add_heading(
      "INDIRECT ASSESSMENT: STAKEHOLDER INTERVIEW NOTES (DE-IDENTIFIED)", level=1
  )
  doc.add_paragraph(
      "Client ID: [CLIENT_ID] | Target Cohort:"
      f" {cohort_meta[cohort_key]['title']}
Informants: Parent, Lead"
      " Therapist / Educator, RBT Supervisor
"
  )
  if cohort_key == "g2":
    doc.add_heading(f"1. Stakeholder Interview — {tb1}", level=2)
    doc.add_paragraph(
        f"Summary ({tb1}): Teachers and aides report tantrums (loud yelling,"
        " foot-stomping, slapping/pushing surfaces) with physical aggression"
        " when preferred activities are denied or transitions occur abruptly."
        " Hypothesized primary function: Tangible access (with transition escape)."
    )
    doc.add_heading(f"2. Stakeholder Interview — {tb2}", level=2)
    doc.add_paragraph(
        f"Summary ({tb2}): Staff observe task refusal and elopement from the"
        " learning area/classroom when multi-step or lengthy independent"
        " academic tasks are assigned without chunking or break options."
        " Hypothesized primary function: Escape from academic demands."
    )
    doc.add_heading(f"3. Stakeholder Interview — {tb3}", level=2)
    doc.add_paragraph(
        f"Summary ({tb3}): Staff report high-volume screaming during auditory"
        " overload in loud school environments such as the cafeteria, gym,"
        " and assemblies. Hypothesized primary function: Sensory regulation /"
        " escape from loud noise."
    )
  elif cohort_key == "g1":
    doc.add_heading(f"1. Stakeholder Interview — {tb1}", level=2)
    doc.add_paragraph(
        f"Summary ({tb1}): Caregivers and RBTs report head banging and wrist"
        " biting during fine-motor/table demands and preferred item removal,"
        " often followed by task pause. Hypothesized primary function: Task"
        " escape (with secondary physical discomfort relief)."
    )
    doc.add_heading(f"2. Stakeholder Interview — {tb2}", level=2)
    doc.add_paragraph(
        f"Summary ({tb2}): Staff report loud screaming and face-slapping"
        " triggered by sudden loud or high-pitched noise (e.g., blender, timer)."
        " Hypothesized primary function: Sensory regulation / escape from"
        " auditory overstimulation."
    )
    doc.add_heading(f"3. Stakeholder Interview — {tb3}", level=2)
    doc.add_paragraph(
        f"Summary ({tb3}): Informants note floor dropping and pica/mouthing"
        " attempts during cleanup and outdoor-to-indoor transitions."
        " Hypothesized primary function: Transition escape & delay (with"
        " outdoor tangible access)."
    )
  else:
    doc.add_heading(f"1. Stakeholder Interview — {tb1}", level=2)
    doc.add_paragraph(
        f"Summary ({tb1}): Job coaches and residential staff report vocational"
        " task refusal with loud verbal threats/slamming when quotas or work"
        " demands increase. Hypothesized primary function: Escape from"
        " vocational demands."
    )
    doc.add_heading(f"2. Stakeholder Interview — {tb2}", level=2)
    doc.add_paragraph(
        f"Summary ({tb2}): Staff report throwing/sweeping objects during"
        " disputes over shared media/spaces (e.g., TV remote). Hypothesized"
        " primary function: Tangible access / control over shared media."
    )
    doc.add_heading(f"3. Stakeholder Interview — {tb3}", level=2)
    doc.add_paragraph(
        f"Summary ({tb3}): Program staff report elopement from the facility"
        " perimeter during unexpected alarms or abrupt schedule shifts."
        " Hypothesized primary function: Escape from sudden sensory/alarm"
        " overload."
    )
  doc.add_heading("4. Historical Reinforcement & Replacement Preferences", level=2)
  doc.add_paragraph(
      "Summary: Historical patterns show problem behaviors often gained"
      " immediate escape or tangibles; stakeholders favor proactive visual"
      " supports, FCT/FERB cards, and calm reinforcement."
  )
  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio

arget Behavior #2: 4-Week Frequency Trend (Zigzag)"
  else:
    freqs = [10, 14, 18, 22]
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
  ax.set_xlabel("Observation Period (4 Weeks)", fontsize=8)
  ax.set_ylabel("Episodes / Total", fontsize=8)
  ax.grid(True, linestyle="--", alpha=0.5)
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
      " task transitions, fine-motor academic demands, and high-pitch sensory"
      " noise environments. Parent notes consistent triggers at home."
  )
  doc.add_heading("2. Interview Note Segment 2 (Educator / RBT Perspective)", level=2)
  doc.add_paragraph(
      "Summary: Staff observe functional escape behaviors when demands are"
      " presented rapidly without priming or visual schedules."
  )
  doc.add_heading("3. Interview Note Segment 3 (Peer / Environment Context)", level=2)
  doc.add_paragraph(
      "Summary: Environmental changes and loud auditory stimuli consistently"
      " exacerbate emotional distress and regulatory difficulties."
  )
  doc.add_heading("4. Interview Note Segment 4 (Historical Reinforcement Patterns)", level=2)
  doc.add_paragraph(
      "Summary: Historical data indicates previous accidental reinforcement of"
      " avoidance through task removal."
  )
  doc.add_heading("5. Interview Note Segment 5 (Intervention Preferences)", level=2)
  doc.add_paragraph(
      "Summary: Stakeholders strongly prefer positive replacement strategies"
      " utilizing visual supports and communication devices."
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
                    "Frequency: 2-4 episodes/day. Duration: 30s - 3min."
                    " Intensity: Moderate."
                ),
                "dimensions_zh": (
                    "频率：每天 2-4 次。持续时间：30秒 至 3分钟。强度：中度。"
                ),
                "dimensions_es": (
                    "Frecuencia: 2-4 episodios/día. Duración: 30s - 3min."
                    " Intensidad: Moderada."
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
                    "65% Direct ABC Data (appliance noise trigger) + 25%"
                    " Indirect RBT Interview (auditory aversion) + 10% QABF"
                    " (Sensory score 13/15)."
                ),
                "triangulation_zh": (
                    "65% 直接 ABC 数据（电器噪音触发） + 25% 间接 RBT 访谈（听觉敏感过度） + 10%"
                    " QABF 结果（感官得分 13/15）。"
                ),
                "triangulation_es": (
                    "65% Datos ABC + 25% Entrevista RBT + 10% QABF (Sensorial"
                    " 13/15)."
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
                    "Frequency: 1-3 episodes/day. Duration: 1-5min. Intensity:"
                    " Moderate."
                ),
                "dimensions_zh": (
                    "频率：每天 1-3 次。持续时间：1 至 5分钟。强度：中度。"
                ),
                "dimensions_es": (
                    "Frecuencia: 1-3 episodios/día. Duración: 1-5min. Intensidad:"
                    " Moderada."
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
                    "65% Direct ABC Data (playground cleanup trigger) + 25%"
                    " Indirect Teacher Interview (transition delay) + 10% QABF"
                    " (Escape score 12/15)."
                ),
                "triangulation_zh": (
                    "65% 直接 ABC 数据（操场清理触发） + 25% 间接教师访谈（转换延缓） + 10%"
                    " QABF 结果（逃避得分 12/15）。"
                ),
                "triangulation_es": (
                    "65% Datos ABC + 25% Entrevista + 10% QABF (Escape 12/15)."
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
            "• Aligned with IDEA IEP requirements and PBIS Multi-Tiered Support Systems.",
            "• Targets academic task engagement, self-advocacy, and emotional self-regulation.",
            "• Emphasizes replacement behaviors integrated into classroom routines.",
            "• 📐 <strong>Triangulation Algorithm:</strong> 65% Direct ABC + 25% Interview + 10% QABF",
        ],
        "behaviors": [
            {
                "name": "1. Tantrum and Physical Aggression",
                "name_zh": "1. 发脾气与身体攻击",
                "name_es": "1. Berrinche y agresión física",
                "def": "Episodes of frustration-related tantrum behavior—including loud yelling at others, foot-stomping, or slapping/pushing desks or objects—co-occurring with physical aggression such as hitting, pushing, or grabbing peers/staff. Onset: initial yelling, hitting, pushing, or forceful object strike; Offset: 15 consecutive seconds without yelling or physical contact/strikes.",
                "def_zh": "与受挫相关的发脾气行为发作——包括朝他人大喊大叫、跺脚，或拍打/推搡课桌或物品——同时伴随身体攻击，如打人、推人或抓握同伴/教职工。以首次大喊、打人、推人或用力拍击物品为行为开始，以连续 15 秒无大喊且无身体接触/拍击为行为结束。",
                "def_es": "Episodios de berrinche relacionados con la frustración—incluyendo gritar fuerte a otros, golpear con los pies o empujar/golpear escritorios u objetos—junto con agresión física como golpear, empujar o sujetar a compañeros o al personal.",
                "ex": "Yelling loudly at teacher, stomping foot, and pushing chair or hitting/slapping surface when denied preferred activity or transition.",
                "ex_zh": "当偏好活动或转换被拒绝时，朝教师大声喊叫、跺脚，并推倒椅子或拍打/击打表面。",
                "ex_es": "Gritar fuerte al maestro, golpear con los pies y empujar una silla o golpear una superficie al negarse una actividad preferida.",
                "non_ex": "Raising hand or speaking calmly to ask a question or request a change in activity.",
                "non_ex_zh": "举手或平静地说话以提出问题或请求调整活动。",
                "non_ex_es": "Levantar la mano o hablar con calma para hacer una pregunta o solicitar un cambio de actividad.",
                "dimensions": "Frequency: 2-4 times per school day. Duration: 30 seconds to 3 minutes per episode. Intensity: Moderate.",
                "dimensions_zh": "频率：每个学校日 2-4 次。持续时间：每次 30 秒至 3 分钟。强度：中度。",
                "dimensions_es": "Frecuencia: 2-4 veces por día escolar. Duración: 30 segundos a 3 minutos. Intensidad: Moderada.",
                "triggers": "Setting Events: Schedule transitions, denied preferred items/activities, high frustration.\nImmediate Triggers: Denial of preferred access, abrupt transition demand, or peer/staff friction.",
                "triggers_zh": "背景事件：日程转换、偏好物品/活动受限、高度受挫。\n直接触发因素：拒绝偏好获取、突如其来的转换要求，或同伴/师生摩擦。",
                "triggers_es": "Eventos de contexto: Transiciones de horario, acceso restringido a preferidos, frustración.\nDesencadenantes inmediatos: Negativa de acceso, demanda abrupta o fricción con pares/personal.",
                "consequences": "Staff clears nearby peers, maintains neutral non-reinforcing response, safely blocks physical strikes/pushes, and prompts functional request once calm.",
                "consequences_zh": "教职工疏散周边同伴，保持中立无强化回应，安全阻挡打人/推搡动作，并在平静后提示功能性请求。",
                "consequences_es": "El personal despeja a compañeros cercanos, mantiene respuesta neutral, bloquea golpes de forma segura e indica una solicitud funcional al calmarse.",
                "qabf_summary": "Tangible Access: 14/15 | Task Escape: 11/15 | Attention: 4/15 | Sensory: 2/15 | Physical: 1/15",
                "qabf_summary_zh": "获得物质: 14/15 | 逃避任务: 11/15 | 社交关注: 4/15 | 感官刺激: 2/15 | 身体不适: 1/15",
                "qabf_summary_es": "Acceso a tangible: 14/15 | Escape: 11/15 | Atención: 4/15 | Sensorial: 2/15 | Físico: 1/15",
                "triangulation": "65% Direct ABC Data + 25% Indirect IEP Interview + 10% QABF (Tangible score 14/15; Escape score 11/15).",
                "triangulation_zh": "65% 直接 ABC 数据 + 25% 间接 IEP 访谈 + 10% QABF 结果（物质得分 14/15；逃避得分 11/15）。",
                "triangulation_es": "65% Datos ABC + 25% Entrevista IEP + 10% QABF (Tangible 14/15; Escape 11/15).",
                "hypothesis": "Primary Function: Access to Tangibles / Preferred Activities (with secondary Escape from abrupt transitions).",
                "hypothesis_zh": "核心行为功能：获取物品/偏好活动（次要功能为逃避突如其来的转换）。",
                "hypothesis_es": "Función principal: Acceso a tangibles / actividades preferidas (con escape secundario de transiciones abruptas).",
                "ferb": "Use a 'Wait / Later' or 'More Time' card/phrase to negotiate access or transition timing calmly.",
                "ferb_zh": "使用“稍等/稍后”或“再多一点时间”卡片/短语，平静地协商获取或转换时机。",
                "ferb_es": "Usar una tarjeta o frase de 'Esperar / Más tarde' o 'Más tiempo' para negociar con calma.",
            },
            {
                "name": "2. Task Refusal and Elopement",
                "name_zh": "2. 任务拒绝与擅自离开区域",
                "name_es": "2. Rechazo de tareas y fuga del área",
                "def": "Refusing to initiate or complete assigned academic tasks (e.g., pushing materials away, crossing arms, saying 'no') accompanied or followed by leaving the assigned learning area or classroom without permission for >5 seconds. Onset: task refusal or crossing designated learning-area boundary; Offset: returning to assigned learning area.",
                "def_zh": "拒绝开始或完成指定的学业任务（如推开材料、双臂交叉、说“不”），伴随或随后未经允许离开指定的学习区域或教室超过 5 秒。以任务拒绝或跨越指定学习区边界为行为开始，以返回指定学习区为行为结束。",
                "def_es": "Negarse a iniciar o completar tareas académicas asignadas (p. ej., apartar materiales, cruzar los brazos, decir 'no') acompañado o seguido de salir del área de aprendizaje o aula sin permiso durante >5 segundos.",
                "ex": "Pushing math/lab worksheet away, refusing to begin, and walking or exiting out of the assigned learning area/classroom during independent work.",
                "ex_zh": "在独立作业期间推开数学/实验工作表、拒绝开始，并在指定学习区/教室内走动或走出该区域。",
                "ex_es": "Apartar la hoja de matemáticas/laboratorio, negarse a comenzar y salir del área de aprendizaje asignada durante el trabajo independiente.",
                "non_ex": "Asking for help or requesting a short scheduled break using a card or word while staying in the designated area.",
                "non_ex_zh": "留在指定区域内，使用卡片或词语请求帮助或申请短暂休息。",
                "non_ex_es": "Pedir ayuda o solicitar un descanso breve usando una tarjeta o palabra permaneciendo en el área designada.",
                "dimensions": "Frequency: 3-5 times per school day. Duration: 1-5 minutes per instance. Intensity: Low to Moderate.",
                "dimensions_zh": "频率：每个学校日 3-5 次。持续时间：每次 1-5 分钟。强度：低至中度。",
                "dimensions_es": "Frecuencia: 3-5 veces por día escolar. Duración: 1-5 minutos. Intensidad: Baja a Moderada.",
                "triggers": "Setting Events: Multi-step or lengthy independent academic assignments.\nImmediate Triggers: Presentation of non-preferred academic worksheet or writing task.",
                "triggers_zh": "背景事件：多步骤或篇幅较长的独立学业作业。\n直接触发因素：呈现非偏好的学业工作表或书写任务。",
                "triggers_es": "Eventos de contexto: Tareas académicas independientes largas o de varios pasos.\nDesencadenantes inmediatos: Presentación de hoja de trabajo o tarea de escritura no preferida.",
                "consequences": "Staff prompts 'Break' or 'Help' card, blocks unsafe exit without escalating, and guides student back to task demand after brief structured break.",
                "consequences_zh": "教职工提示使用“休息”或“帮助”卡片，在不激化情绪的前提下阻挡不安全离开，并在短暂结构化休息后引导返回任务。",
                "consequences_es": "El personal indica la tarjeta 'Descanso' o 'Ayuda', bloquea salidas inseguras sin escalar y guía de vuelta a la tarea tras un descanso breve.",
                "qabf_summary": "Task Escape: 15/15 | Attention: 4/15 | Tangible: 2/15 | Sensory: 1/15 | Physical: 0/15",
                "qabf_summary_zh": "逃避任务: 15/15 | 社交关注: 4/15 | 获得物质: 2/15 | 感官刺激: 1/15 | 身体不适: 0/15",
                "qabf_summary_es": "Escape: 15/15 | Atención: 4/15 | Tangible: 2/15 | Sensorial: 1/15 | Físico: 0/15",
                "triangulation": "65% Direct ABC Data + 25% Indirect IEP Interview + 10% QABF (Escape score 15/15).",
                "triangulation_zh": "65% 直接 ABC 数据 + 25% 间接 IEP 访谈 + 10% QABF 结果（逃避得分 15/15）。",
                "triangulation_es": "65% Datos ABC + 25% Entrevista IEP + 10% QABF (Escape 15/15).",
                "hypothesis": "Primary Function: Escape from Difficult or Non-Preferred Academic Demands.",
                "hypothesis_zh": "核心行为功能：逃避困难或非偏好的学业要求。",
                "hypothesis_es": "Función principal: Escape de demandas académicas difíciles o no preferidas.",
                "ferb": "Hand 'Break' or 'Help' card or request a chunked task option to pause or modify academic demand.",
                "ferb_zh": "递交“休息”或“帮助”卡片，或请求拆分任务选项以暂停或调整学业要求。",
                "ferb_es": "Entregar tarjeta de 'Descanso' o 'Ayuda' o solicitar una opción de tarea dividida.",
            },
            {
                "name": "3. Screaming (Auditory Overload)",
                "name_zh": "3. 听觉过载下的高分贝尖叫",
                "name_es": "3. Gritos (Sobrecarga auditiva)",
                "def": "High-pitched or high-volume vocal screaming (>75-80 dB) occurring during or immediately following loud or abrupt auditory stimuli (e.g., cafeteria noise, bells, gym/assembly noise), often with hands near ears. Onset: initial high-volume scream; Offset: 10 consecutive seconds of quiet vocalizations.",
                "def_zh": "在处于或紧随高分贝或突发性听觉刺激（如食堂噪音、铃声、体育馆/集会噪音）之后出现的高音调或高音量尖叫（>75-80 dB），常伴有双手靠近耳朵。以首次高音量尖叫为行为开始，以连续 10 秒安静发声为行为结束。",
                "def_es": "Gritos vocales agudos o de alto volumen (>75-80 dB) que ocurren durante o inmediatamente después de estímulos auditivos fuertes o repentinos (p. ej., ruido de cafetería, timbres, gimnasio/asamblea).",
                "ex": "Screaming loudly in the cafeteria when trays clatter and ambient noise spikes.",
                "ex_zh": "在食堂餐盘碰撞、环境噪音骤增时大声尖叫。",
                "ex_es": "Gritar fuerte en la cafetería cuando chocan las bandejas y sube el ruido ambiental.",
                "non_ex": "Cheering or speaking at normal volume during structured games or quiet conversation.",
                "non_ex_zh": "在结构化游戏或安静交谈期间欢呼或以正常音量说话。",
                "non_ex_es": "Animar o hablar a volumen normal durante juegos estructurados o conversación tranquila.",
                "dimensions": "Frequency: 1-3 times weekly. Duration: 15 seconds to 2 minutes per episode. Intensity: Moderate to High (loud vocal volume).",
                "dimensions_zh": "频率：每周 1-3 次。持续时间：每次 15 秒至 2 分钟。强度：中度至高度（高音量发声）。",
                "dimensions_es": "Frecuencia: 1-3 veces por semana. Duración: 15 segundos a 2 minutos. Intensidad: Moderada a Alta.",
                "triggers": "Setting Events: Loud school environments (cafeteria, assemblies, gymnasium, passing periods).\nImmediate Triggers: Sudden loud noise, clattering trays, or high ambient sound levels.",
                "triggers_zh": "背景事件：嘈杂的校园环境（食堂、集会、体育馆、课间走动时段）。\n直接触发因素：突发巨响、餐盘碰撞声或高环境噪音水平。",
                "triggers_es": "Eventos de contexto: Entornos escolares ruidosos (cafetería, asambleas, gimnasio).\nDesencadenantes inmediatos: Ruido fuerte repentino, bandejas o alto nivel de ruido ambiental.",
                "consequences": "Staff offers noise-canceling headphones and guides student to a quiet sensory corner or space.",
                "consequences_zh": "教职工提供降噪耳机，并引导学生前往安静的感官角落或空间。",
                "consequences_es": "El personal ofrece auriculares con cancelación de ruido y guía al estudiante a un rincón sensorial tranquilo.",
                "qabf_summary": "Sensory Automatic: 14/15 | Escape: 10/15 | Attention: 1/15",
                "qabf_summary_zh": "感官自动强化: 14/15 | 逃避任务: 10/15 | 社交关注: 1/15",
                "qabf_summary_es": "Sensorial automático: 14/15 | Escape: 10/15 | Atención: 1/15",
                "triangulation": "65% Direct ABC Data + 25% Indirect Staff Interview + 10% QABF (Sensory score 14/15).",
                "triangulation_zh": "65% 直接 ABC 数据 + 25% 间接教职工访谈 + 10% QABF 结果（感官得分 14/15）。",
                "triangulation_es": "65% Datos ABC + 25% Entrevista del personal + 10% QABF (Sensorial 14/15).",
                "hypothesis": "Primary Function: Automatic Sensory Regulation / Escape from Avversive Auditory Stimulation.",
                "hypothesis_zh": "核心行为功能：自动感官调节/逃避令人不适的听觉刺激。",
                "hypothesis_es": "Función principal: Regulación sensorial automática / Escape de estimulación auditiva aversiva.",
                "ferb": "Put on noise-canceling headphones or show a 'Quiet Space / Headphones' card when noise rises.",
                "ferb_zh": "当噪音升高时戴上降噪耳机，或出示“安静空间/耳机”卡片。",
                "ferb_es": "Ponerse auriculares con cancelación de ruido o mostrar una tarjeta de 'Espacio tranquilo / Auriculares'.",
            },
        ],
        "strengths": "Excellent visual-spatial abilities, enthusiastic about technology and drawing.",
        "strengths_zh": "具备出色的视觉空间能力，对科技和绘画抱有极高热情。",
        "strengths_es": "Excelentes habilidades visoespaciales, entusiasta de la tecnología y el dibujo.",
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
            "• Designed for Medicaid HCBS Waiver adult day programs and community living.",
            "• Focuses on person-centered planning, independence, and vocational endurance.",
            "• Emphasizes self-management protocols and respectful adult communication.",
            "• 📐 <strong>Triangulation Algorithm:</strong> 65% Direct ABC + 25% Interview + 10% QABF",
        ],
        "behaviors": [
            {
                "name": "1. Vocational Task Refusal & Verbal Aggression",
                "name_zh": "1. 拒绝职业组装任务与言语攻击",
                "name_es": "1. Rechazo de tareas vocacionales y agresión verbal",
                "def": "Refusing assembly or sorting demands accompanied by loud vocal threats (>75 dB) or pushing work materials away. Onset: vocal outburst or material push; Offset: 3 minutes of quiet task engagement.",
                "def_zh": "拒绝组装或分类任务，并伴随大声言语威胁（>75 dB）或推开工作材料的行为。以言语发作或推开材料为行为开始，以连续 3 分钟安静参与任务为行为结束。",
                "def_es": "Rechazar demandas de ensamblaje acompañado de amenazas vocales fuertes (>75 dB) o empujar los materiales de trabajo.",
                "ex": "Shouting 'No way!', slamming assembly boxes on desk when quota is raised.",
                "ex_zh": "当提高工作配额时，大叫“绝不可能！”并将组装盒重重摔在桌上。",
                "ex_es": "Gritar '¡De ninguna manera!' y golpear cajas sobre la mesa.",
                "non_ex": "Verbally requesting a 5-minute break in a normal tone.",
                "non_ex_zh": "用正常音量和语调口头提出“想要休息5分钟”。",
                "non_ex_es": "Solicitar verbalmente un descanso de 5 minutos en tono normal.",
                "dimensions": "Frequency: 1-2 times weekly. Duration: 5-10 minutes. Intensity: Moderate.",
                "dimensions_zh": "频率：每周 1-2 次。持续时间：5-10 分钟。强度：中度。",
                "dimensions_es": "Frecuencia: 1-2 veces por semana. Duración: 5-10 min. Intensidad: Moderada.",
                "triggers": "Setting Events: Unfamiliar staff.\nImmediate Triggers: Direct instructions to complete vocational assembly quota.",
                "triggers_zh": "背景事件：不熟悉的工作人员。\n直接触发因素：要求完成职业组装配额的直接指令。",
                "triggers_es": "Eventos de contexto: Personal no familiar.\nDesencadenante: Instrucción directa de ensamblaje.",
                "consequences": "DSP offers choice board, demand temporarily paused.",
                "consequences_zh": "直属支持人员（DSP）提供选择板，任务要求被暂时暂停。",
                "consequences_es": "El personal DSP ofrece panel de opciones, demanda pausada.",
                "qabf_summary": "Task Escape: 13/15 | Attention: 4/15 | Tangible: 3/15 | Sensory: 1/15 | Physical: 1/15",
                "qabf_summary_zh": "逃避任务: 13/15 | 社交关注: 4/15 | 获得物质: 3/15 | 感官刺激: 1/15 | 身体不适: 1/15",
                "qabf_summary_es": "Escape: 13/15 | Atención: 4/15 | Tangible: 3/15 | Sensorial: 1/15 | Físico: 1/15",
                "triangulation": "65% Direct ABC Data + 25% Indirect Vocational Interview + 10% QABF (Escape score 13/15).",
                "triangulation_zh": "65% 直接 ABC 数据 + 25% 间接职业能力访谈 + 10% QABF 结果（逃避得分 13/15）。",
                "triangulation_es": "65% Datos ABC + 25% Entrevista + 10% QABF (Escape 13/15).",
                "hypothesis": "Primary Function: Escape from Vocational Assembly Demands.",
                "hypothesis_zh": "核心行为功能：逃避职业组装工作要求。",
                "hypothesis_es": "Función principal: Escape de demandas vocacionales.",
                "ferb": "Verbally request '5-minute break, please' using self-advocacy phrase card.",
                "ferb_zh": "使用自我倡导短语卡口头表达：“请给我 5 分钟休息时间”。",
                "ferb_es": "Solicitar verbalmente 'Descanso de 5 minutos, por favor'.",
            },
            {
                "name": "2. Property Disruption During Peer Conflict",
                "name_zh": "2. 同伴冲突期间的物品破坏行为",
                "name_es": "2. Alteración de la propiedad durante conflicto con pares",
                "def": "Throwing objects or sweeping items off surfaces when disputes arise over shared community spaces or items. Onset: grabbing/throwing object; Offset: 2 minutes of calm.",
                "def_zh": "当在共享社区空间或物品产生争议时，投掷物品或将物品从表面扫落。以抓取/投掷物品为行为开始，以持续 2 分钟平静为行为结束。",
                "def_es": "Lojar objetos o barrer elementos de las superficies cuando surgen disputas.",
                "ex": "Throwing couch cushions when peer takes remote control.",
                "ex_zh": "当同伴拿走遥控器时扔掉沙发布艺垫。",
                "ex_es": "Lanzar cojines del sofá cuando un compañero toma el control remoto.",
                "non_ex": "Asking staff to mediate peer disagreement calmly.",
                "non_ex_zh": "平静地请求工作人员调解同伴意见分歧。",
                "non_ex_es": "Pedir al personal que medie en el desacuerdo con calma.",
                "dimensions": "Frequency: 1-2 times weekly. Duration: 2-5 minutes. Intensity: Moderate.",
                "dimensions_zh": "频率：每周 1-2 次。持续时间：2-5 分钟。强度：中度。",
                "dimensions_es": "Frecuencia: 1-2 veces por semana. Duración: 2-5 min. Intensidad: Moderada.",
                "triggers": "Setting Events: Shared living spaces, competing for media items.",
                "triggers_zh": "背景事件：共享生活空间、争夺娱乐设备。",
                "triggers_es": "Eventos de contexto: Espacios compartidos, competencia por medios.",
                "consequences": "Staff intervenes, separates individuals, offers mediation.",
                "consequences_zh": "工作人员干预，分隔双方并提供调解。",
                "consequences_es": "El personal interviene, separa a las personas y ofrece mediación.",
                "qabf_summary": "Tangible Access: 14/15 | Escape: 8/15 | Attention: 5/15",
                "qabf_summary_zh": "获得物质: 14/15 | 逃避任务: 8/15 | 社交关注: 5/15",
                "qabf_summary_es": "Acceso a tangible: 14/15 | Escape: 8/15 | Atención: 5/15",
                "triangulation": "65% ABC Data + 25% DSP Interview + 10% QABF.",
                "triangulation_zh": "65% ABC 数据 + 25% DSP 访谈 + 10% QABF。",
                "triangulation_es": "65% Datos ABC + 25% Entrevista DSP + 10% QABF.",
                "hypothesis": "Primary Function: Access to Tangible / Control over Shared Media.",
                "hypothesis_zh": "核心行为功能：获取物质/控制共享媒介。",
                "hypothesis_es": "Función principal: Acceso a tangible / Control de medios.",
                "ferb": "Use communication board to negotiate turn-taking schedule with peers.",
                "ferb_zh": "使用沟通板与同伴协商轮流使用时间表。",
                "ferb_es": "Usar tablero de comunicación para negociar turnos con pares.",
            },
            {
                "name": "3. Elopement from Facility During Schedule Shifts",
                "name_zh": "3. 日程调整期间擅离社区机构",
                "name_es": "3. Fuga de la instalación durante cambios de horario",
                "def": "Leaving designated day program perimeter without staff notification during unplanned alarms or schedule transitions. Onset: crossing perimeter exit; Offset: staff intervention.",
                "def_zh": "在未预料的警报或日程转换期间，未经工作人员通知擅自离开指定的日间项目边界。以跨越边界出口为行为开始，以工作人员干预为行为结束。",
                "def_es": "Abandonar el perímetro designado del programa diurno sin notificar al personal.",
                "ex": "Running out door during unexpected fire drill alarm.",
                "ex_zh": "在突发消防演习警报期间跑出门外。",
                "ex_es": "Correr hacia afuera durante una alarma de simulacro de incendio.",
                "non_ex": "Walking in orderly line with group during scheduled drills.",
                "non_ex_zh": "在计划演习期间随队伍有序排队行走。",
                "non_ex_es": "Caminar en fila ordenada con el grupo durante simulacros programados.",
                "dimensions": "Frequency: 1 time monthly. Duration: 10-15 minutes. Intensity: Moderate to High.",
                "dimensions_zh": "频率：每月 1 次。持续时间：10-15 分钟。强度：中度至高度。",
                "dimensions_es": "Frecuencia: 1 vez al mes. Duración: 10-15 min. Intensidad: Moderada a Alta.",
                "triggers": "Setting Events: Sudden loud alarms, unexpected transition triggers.",
                "triggers_zh": "背景事件：突发高分贝警报、意外转换触发因素。",
                "triggers_es": "Eventos de contexto: Alarmas fuertes repentinas, transiciones inesperadas.",
                "consequences": "Staff escorts client safely back and reviews visual schedule.",
                "consequences_zh": "工作人员护送客户安全返回并复习视觉日程表。",
                "consequences_es": "El personal acompaña al cliente de regreso y revisa el horario visual.",
                "qabf_summary": "Escape: 13/15 | Sensory/Anxiety: 11/15 | Attention: 2/15",
                "qabf_summary_zh": "逃避任务: 13/15 | 感官/焦虑: 11/15 | 社交关注: 2/15",
                "qabf_summary_es": "Escape: 13/15 | Sensorial/Ansiedad: 11/15 | Atención: 2/15",
                "triangulation": "65% ABC Data + 25% Staff Interview + 10% QABF.",
                "triangulation_zh": "65% ABC 数据 + 25% 工作人员访谈 + 10% QABF。",
                "triangulation_es": "65% Datos ABC + 25% Entrevista + 10% QABF.",
                "hypothesis": "Primary Function: Escape from Sudden Sensory/Alarm Overload.",
                "hypothesis_zh": "核心行为功能：逃避突发感官/警报超载。",
                "hypothesis_es": "Función principal: Escape de sobrecarga sensorial/alarma repentina.",
                "ferb": "Request staff accompaniment or quiet transition space using card.",
                "ferb_zh": "使用卡片请求工作人员陪同或安静的转换空间。",
                "ferb_es": "Solicitar acompañamiento del personal o espacio tranquilo con tarjeta.",
            },
        ],
        "strengths": "High independence in personal self-care.",
        "strengths_zh": "在个人日常生活自理方面具备极高独立性。",
        "strengths_es": "Alta independencia en el cuidado personal.",
        "history": "Participates in Adult Day Vocational Services under Medicaid HCBS Waiver.",
        "history_zh": "在 Medicaid HCBS 豁免计划下参与成人日间职业干预服务。",
        "history_es": "Participa en servicios vocacionales para adultos bajo exención Medicaid HCBS.",
    },
}

# ==========================================
# 3. Main Interface & Security Header
# ==========================================
st.markdown(
    """
    <div class="hipaa-banner">
        <div class="hipaa-title">🛡️ 100% HIPAA COMPLIANT & BACB ETHICS-ALIGNED LOCAL PROCESSING</div>
        <div class="hipaa-body">
            <b>This tool strictly complies with HIPAA privacy regulations and aligns with the</b> <b>BACB Ethics Code (Section 2.11 Maintaining Confidentiality)</b>. <b>All data parsing, analysis, and document formulation occur </b> <b>100% locally within your active browser session memory</b>.
        </div>
    </div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    "<div class='main-title'>🧩 BIP Engine - BCBA Clinical FBA & BIP Draft Formulation Tool"
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

# Phase 1 vs. Phase 2 Expectation Warning placed right above upload buttons
st.warning(
    "⚠️ **Phase 1 (Demo Version) vs. Phase 2 (Local Deployment Roadmap)**\n\n"
    "• **Current Phase 1 (Demo):** You are evaluating the demo version. **Strictly use Mock Data provided below only.** Do not upload any files containing real client health information, PII (Personally Identifiable Information), or PHI (Protected Health Information).\n\n"
    "• **Phase 2 Blueprint (Local Landing):** Future production will run 100% locally on the BCBA's desktop machine. Workflow: Upload raw client notes locally -> Web app automatically scans, desensitizes, and replaces PHI with secure placeholders (e.g., `[CLIENT_NAME]`, `[CLIENT_ID]`) -> BCBA downloads the draft and finalizes instantly using `Ctrl+H` text replacement."
)

col_input1, col_input2, col_input3 = st.columns([1, 1, 1])

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


def build_compact_demographics_table(doc, c_meta, lang_mode, template_choice="School-Based IEP"):
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
            "Template Type (模板类型)",
            template_choice,
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
            "Template Type (Plantilla)",
            template_choice,
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
            "Template Type",
            template_choice,
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


def _add_apa_entry(doc, segments):
  p = doc.add_paragraph()
  p.paragraph_format.left_indent = Inches(0.5)
  p.paragraph_format.first_line_indent = Inches(-0.5)
  p.paragraph_format.space_after = Pt(5)
  for text, is_italic in segments:
    run = p.add_run(text)
    if is_italic:
      run.italic = True
  return p


def _add_ref_section_label(doc, label):
  p = doc.add_paragraph()
  p.paragraph_format.space_before = Pt(6)
  p.paragraph_format.space_after = Pt(2)
  r = p.add_run(label)
  r.bold = True
  r.font.color.rgb = RGBColor(31, 78, 120)


def append_fba_evidence_references(doc, lang_mode=None):
  h = doc.add_heading(level=1)
  h.add_run("References (FBA)")

  # Section 2: Data Sources & Triangulation Methodology (65% ABC, 25% Interview, 10% QABF)
  _add_ref_section_label(
      doc,
      "Section 2 — Data Sources & Triangulation Methodology (65% ABC / 25% Interview / 10% QABF):",
  )
  _add_apa_entry(doc, [
      ("Cooper, J. O., Heron, T. E., & Heward, W. L. (2020). ", False),
      ("Applied behavior analysis", True),
      (" (3rd ed.). Pearson.", False),
  ])
  _add_apa_entry(doc, [
      ("Matson, J. L., & Vollmer, T. R. (1995). ", False),
      ("Questions About Behavioral Function (QABF) user's guide", True),
      (". Disability Associates.", False),
  ])

  # Section 4: Individual Target Behavior Functional Analyses
  _add_ref_section_label(
      doc,
      "Section 4 — Individual Target Behavior Functional Analyses (Definitions, Triggers, Consequences & Hypothesized Function):",
  )
  _add_apa_entry(doc, [
      (
          "Iwata, B. A., Dorsey, M. F., Slifer, K. J., Bauman, K. E., & Richman,"
          " G. S. (1994). Toward a functional analysis of self-injury. ",
          False,
      ),
      ("Journal of Applied Behavior Analysis", True),
      (", 27(2), 197–209. (Reprinted from ", False),
      ("Journal of Applied Behavior Analysis", True),
      (", 1982, 15, 3–21).", False),
  ])
  _add_apa_entry(doc, [
      (
          "O'Neill, R. E., Albin, R. W., Storey, K., Horner, R. H., & Sprague,"
          " J. R. (2015). ",
          False,
      ),
      (
          "Functional assessment and program development for problem behavior:"
          " A practical handbook",
          True,
      ),
      (" (3rd ed.). Cengage Learning.", False),
  ])

  # Section 5: Synthesis & Personalised Clinical Recommendations
  _add_ref_section_label(
      doc,
      "Section 5 — Synthesis & Personalised Clinical Recommendations:",
  )
  _add_apa_entry(doc, [
      (
          "Horner, R. H., Sugai, G., & Anderson, C. M. (2010). Examining the"
          " evidence base for school-wide positive behavior support. ",
          False,
      ),
      ("Focus on Exceptional Children", True),
      (", 42(8), 1–14.", False),
  ])
  _add_apa_entry(doc, [
      (
          "Carr, E. G., & Durand, V. M. (1985). Reducing behavior problems"
          " through functional communication training. ",
          False,
      ),
      ("Journal of Applied Behavior Analysis", True),
      (", 18(2), 111–126.", False),
  ])


def append_bip_evidence_references(doc, lang_mode=None):
  h = doc.add_heading(level=1)
  h.add_run("References (BIP)")

  # Section 2: Target Behaviors, Functions & Replacement Skills (FERB)
  _add_ref_section_label(
      doc,
      "Section 2 — Target Behaviors, Functions & Replacement Skills (FERB):",
  )
  _add_apa_entry(doc, [
      ("Cooper, J. O., Heron, T. E., & Heward, W. L. (2020). ", False),
      ("Applied behavior analysis", True),
      (" (3rd ed.). Pearson.", False),
  ])
  _add_apa_entry(doc, [
      (
          "Tiger, J. H., Hanley, G. P., & Bruzek, J. L. (2008). Functional"
          " communication training: A review and practical guide. ",
          False,
      ),
      ("Behavior Analysis in Practice", True),
      (", 1(1), 16–23.", False),
  ])

  # Section 3: Proactive & Antecedent Modifications (Prevention)
  _add_ref_section_label(
      doc,
      "Section 3 — Proactive & Antecedent Modifications (Prevention):",
  )
  _add_apa_entry(doc, [
      (
          "Kern, L., & Clemens, N. H. (2007). Antecedent strategies to promote"
          " appropriate classroom behavior. ",
          False,
      ),
      ("Psychology in the Schools", True),
      (", 44(1), 65–75.", False),
  ])
  _add_apa_entry(doc, [
      ("Miltenberger, R. G. (2016). ", False),
      ("Behavior modification: Principles and procedures", True),
      (" (6th ed.). Cengage Learning.", False),
  ])

  # Section 4: Replacement Behaviors & Functional Communication Training (FCT)
  _add_ref_section_label(
      doc,
      "Section 4 — Replacement Behaviors & Functional Communication Training (FCT):",
  )
  _add_apa_entry(doc, [
      (
          "Carr, E. G., & Durand, V. M. (1985). Reducing behavior problems"
          " through functional communication training. ",
          False,
      ),
      ("Journal of Applied Behavior Analysis", True),
      (", 18(2), 111–126.", False),
  ])
  _add_apa_entry(doc, [
      (
          "Durand, V. M., & Moskowitz, L. (2015). Functional communication"
          " training: Thirty years of treating challenging behavior. ",
          False,
      ),
      ("Topics in Early Childhood Special Education", True),
      (", 34(4), 216–226.", False),
  ])

  # Section 5: Reinforcement Protocols
  _add_ref_section_label(
      doc,
      "Section 5 — Reinforcement Protocols:",
  )
  _add_apa_entry(doc, [
      (
          "Vollmer, T. R., Roane, H. S., Ringdahl, J. E., & Marcus, B. A."
          " (1999). Evaluating grading criteria as reinforcement during"
          " functional communication training. ",
          False,
      ),
      ("Journal of Applied Behavior Analysis", True),
      (", 32(4), 451–466.", False),
  ])
  _add_apa_entry(doc, [
      (
          "Poling, A. (2010). Schedules of reinforcement: The definition and use"
          " of reinforcement schedules. ",
          False,
      ),
      ("Journal of Applied Behavior Analysis", True),
      (", 43(1), 7–25.", False),
  ])

  # Section 6: Reactive Response Protocols & Extinction
  _add_ref_section_label(
      doc,
      "Section 6 — Reactive Response Protocols & Extinction:",
  )
  _add_apa_entry(doc, [
      (
          "Iwata, B. A., Pace, G. M., Cowdery, G. E., & Miltenberger, R. G."
          " (1994). What makes extinction work: An analysis of procedural"
          " form and function. ",
          False,
      ),
      ("Journal of Applied Behavior Analysis", True),
      (", 27(1), 131–144.", False),
  ])
  _add_apa_entry(doc, [
      (
          "Lerman, D. C., & Iwata, B. A. (1996). Developing a technology for"
          " the use of extinction in treating problem behavior: A review and"
          " analysis. ",
          False,
      ),
      ("Journal of Applied Behavior Analysis", True),
      (", 29(3), 345–382.", False),
  ])

  # Section 7: Crisis Safety & Treatment Fidelity
  _add_ref_section_label(
      doc,
      "Section 7 — Crisis Safety & Treatment Fidelity:",
  )
  _add_apa_entry(doc, [
      (
          "McGill, P., Bradshaw, J., & Hughes, A. (2009). Impact of extended"
          " staff training in positive behavior support on re-active strategy"
          " use and behavior of people with severe intellectual disabilities. ",
          False,
      ),
      ("Mental Retardation and Developmental Disabilities Research Reviews", True),
      (", 15(3), 246–254.", False),
  ])
  _add_apa_entry(doc, [
      (
          "Fryling, M. J., Wallace, M. D., & Yassine, J. N. (2012). Impact of"
          " treatment integrity on interventions for problem behavior. ",
          False,
      ),
      ("Journal of Applied Behavior Analysis", True),
      (", 45(2), 449–453.", False),
  ])

  # Section 8: Staff Training and Monitoring
  _add_ref_section_label(
      doc,
      "Section 8 — Staff Training and Monitoring:",
  )
  _add_apa_entry(doc, [
      (
          "Parsons, M. B., Rollyson, J. H., & Reid, D. H. (2012). Evidence-based"
          " staff training: A guide for practitioners. ",
          False,
      ),
      ("Behavior Analysis in Practice", True),
      (", 5(2), 2–11.", False),
  ])
  _add_apa_entry(doc, [
      (
          "Fleming, R. K., & Sulzer-Azaroff, B. (1992). Enforcing treatment"
          " fidelity in the classroom. ",
          False,
      ),
      ("Journal of Applied Behavior Analysis", True),
      (", 25(2), 385–397.", False),
  ])


# ==========================================
# 7. Behavior-by-Behavior Separated FBA Generator
# ==========================================
def generate_exact_fba_doc(cohort_key, lang_choice, behavior_list, template_choice="School-Based IEP"):
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
  build_compact_demographics_table(doc, c_meta, lang_mode, template_choice)

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
          f"Template Format: {template_choice}. Triangulation Algorithm Standard"
          " Applied: 65% Direct ABC Data + 25% Indirect Interview Data + 10%"
          " QABF.\n\n1. Direct ABC Data (65%): Continuous recording across"
          " baseline sessions.\n2. Indirect Assessment (25%): Structured"
          " interviews with stakeholder/informant.\n3. Psychometric Rating"
          " Scale (10%): Behavior-specific QABF (Questions About Behavioral"
          " Function)."
      ),
      (
          "评估方法与三方验证公式"
          if lang_mode == "zh"
          else "Metodología y fórmula de triangulación"
          if lang_mode == "es"
          else None
      ),
      (
          f"模板格式：{template_choice}。应用三方数据加权验证算法：65% Direct ABC Data + 25%"
          " Indirect Interview Data + 10% QABF。\n\n1. 直接 ABC 数据（65%）：基线观察期间连续行为记录。\n2."
          " 间接评估（25%）：与信息提供者的结构化访谈。\n3. 心理测量量表（10%）：针对具体行为的 QABF"
          " 评估问卷。"
          if lang_mode == "zh"
          else f"Plantilla: {template_choice}. Algoritmo de triangulación: 65%"
          " ABC + 25% Entrevista + 10% QABF."
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )
  is_school = "School" in template_choice
  add_bi_item(
      doc,
      "Template Organization & Recommendation Focus",
      (
          "Recommendations are organized around School / Classroom & IEP Alignment"
          " (classroom routines, transitions, teacher/paraeducator prompts, IEP"
          " objectives)."
          if is_school
          else "Recommendations are organized around Clinical / Home & Community"
          " Supports (caregiver routines, home/community settings, structured"
          " teaching / reinforcement contingencies)."
      ),
      (
          "模板组织与建议侧重"
          if lang_mode == "zh"
          else "Organización de plantilla y enfoque"
          if lang_mode == "es"
          else None
      ),
      (
          (
              "建议围绕 School / Classroom & IEP Alignment（课堂常规、过渡环节、教师/助教提示、IEP"
              " 目标）组织。"
              if is_school
              else "建议围绕 Clinical / Home & Community"
              " Supports（照护者常规、家庭/社区场景、结构化教学/强化伴随事件）组织。"
          )
          if lang_mode == "zh"
          else (
              "Recomendaciones organizadas en torno a School / Classroom & IEP"
              " Alignment."
              if is_school
              else "Recomendaciones organizadas en torno a Clinical / Home &"
              " Community Supports."
          )
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
      "4. Individual Target Behavior Functional Analyses (4-Week Trends)",
      (
          "4. 目标行为独立功能分析 (按行为逐项拆解与 4 周专属趋势图)"
          if lang_mode == "zh"
          else "4. Análisis funcional individual por conducta (tendencias de 4 semanas)"
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

    # Insert Dedicated Behavior-Specific Tracking Chart (4 Weeks)
    chart_title_en = f"Figure 4.{idx}. 4-Week Longitudinal Behavior Tracking Trend for Target Behavior #{idx}"
    chart_title_zh = f"图 4.{idx}. 目标行为 #{idx} 4周专属纵向追踪趋势图"
    chart_title_es = f"Figura 4.{idx}. Tendencia de seguimiento de 4 semanas para Conducta #{idx}"
    
    add_bi_heading(
        doc,
        3,
        chart_title_en,
        chart_title_zh if lang_mode == "zh" else chart_title_es if lang_mode == "es" else None
    )
    chart_buf = generate_behavior_tracking_chart(cohort_key, idx - 1)
    doc.add_picture(chart_buf, width=Inches(5.0))
    doc.add_paragraph(
        f"Note: The chart above illustrates the independent 4-week tracking frequency trend specifically recorded for Target Behavior #{idx}."
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
        "Based on the triangulated functional assessment (65% Direct ABC + 25%"
        " Indirect Interview + 10% QABF), the client (Ages 2-5) primarily"
        " demonstrates behaviors maintained by Social Negative Reinforcement"
        " (Task/Transition Escape) and Sensory Automatic Regulation"
        " (Auditory Aversion). Recommendations:\n1. Formulate an NDBI / ESDM"
        " play-based Behavior Intervention Plan (BIP) integrating AAC (BigMack"
        " / PECS 'Break' card) for early functional communication training"
        " (FCT).\n2. Implement antecedent noise-mitigation protocols"
        " (noise-canceling headphones, visual schedule warnings) prior to structured"
        " table tasks and transitions.\n3. Train parents and RBTs on shared"
        " co-regulation strategies to replace self-injurious and sensory"
        " distress behaviors."
    )
    synth_zh = (
        "基于 65% 直接 ABC 观察 + 25% 间接访谈 + 10% QABF 量表的三方交叉验证算法，该 2-5"
        " 岁幼儿的目标行为主要由‘逃避任务/转换（社交负强化）’及‘调节听觉过度刺激（感官自动强化）’所维持。临床建议：\n1."
        " 制定基于 ESDM / NDBI 自然教法的行为干预计划 (BIP)，优先引入 AAC"
        " 语音按键与 PECS‘休息’视觉卡进行早期功能性沟通训练 (FCT)。\n2."
        " 在结构化桌面任务与环节转换前，实施前因降噪与视觉倒计时预警。\n3."
        " 对家长及 RBT 进行情绪共调节 (Co-regulation) 技巧培训，以替代自伤与感官情绪发作。"
    )
    synth_es = (
        "Con base en la triangulación (65% ABC + 25% Entrevista + 10% QABF), el"
        " cliente (2-5 años) presenta conductas mantenidas por Escape de"
        " Tareas y Regulación Sensorial Auditiva. Recomendaciones:\n1. Diseñar"
        " un BIP basado en NDBI/ESDM integrando AAC/PECS para Comunicación"
        " Funcional (FCT).\n2. Implementar protocolos de mitigación de ruido y"
        " avisos visuales antes de transiciones.\n3. Capacitar a padres y RBTs"
        " en co-regulación."
    )
  elif cohort_key == "g2":
    synth_en = (
        "Based on triangulated assessment data (65% ABC Data + 25% IEP Teacher"
        " Interviews + 10% QABF Assessment), the student (Ages 5-21) demonstrates"
        " frustration-related tantrums and physical aggression around transitions/preferred"
        " items, task refusal and elopement from the learning area during academic"
        " demands, and high-volume screaming during auditory overload in loud"
        " environments (e.g., cafeteria, gymnasium).\n\n"
        "Comprehensive Personalized Clinical Recommendations:\n"
        "1. Academic Task Modification & Pacing: Implement visual breakdown"
        " of multi-step assignments, alternating preferred and non-preferred"
        " academic tasks to reduce frustration and build academic endurance.\n"
        "2. Functional Communication & Self-Advocacy (FCT): Integrate"
        " discrete desk cards ('Help Needed', 'Break Request', 'More Time') to"
        " negotiate transitions and demands without escalation.\n"
        "3. Sensory Regulation Support: Provide proactive access to"
        " noise-canceling headphones and designated quiet sensory passes during"
        " high-stimulation school settings (cafeteria, assembly halls, gym).\n"
        "4. PBIS & Multi-Tiered Positive Reinforcement: Utilize positive token"
        " economies and immediate descriptive praise for independent task"
        " engagement and appropriate self-regulation."
    )
    synth_zh = (
        "基于三方交叉验证评估数据（65% ABC 数据 + 25% IEP 教师访谈 + 10% QABF 量表），该学龄期学生（5-21岁）"
        "主要表现出与转换/偏好物品相关的发脾气与身体攻击、学业要求下的任务拒绝与擅离学习区，以及在食堂、体育馆等高噪环境下的听觉过载高分贝尖叫。\n\n"
        "个性化综合临床建议：\n"
        "1. 学业任务调整与节奏控制：将多步骤作业进行视觉拆解，交替安排偏好与非偏好任务，以降低挫败感并培养学业耐受力。\n"
        "2. 功能性沟通与自我倡导 (FCT)：引入桌面提示卡（“需要帮助”、“休息申请”、“再多一点时间”），"
        "在无需升级行为的情况下协商转换与任务要求。\n"
        "3. 感官调节支持：在食堂、礼堂、体育馆等高刺激校园场景，主动提供降噪耳机与指定的安静空间通行证。\n"
        "4. PBIS 多级正向强化：利用代币制与即时描述性表扬，对学生独立参与任务及恰当的自我调节行为给予正向强化。"
    )
    synth_es = (
        "Basado en los datos triangulados (65% ABC + 25% Entrevistas IEP + 10% QABF),"
        " el estudiante (5-21 años) presenta berrinches y agresión física, rechazo de"
        " tareas y fuga del área de aprendizaje, y gritos por sobrecarga auditiva.\n\n"
        "Recomendaciones clínicas personalizadas:\n"
        "1. Modificación de tareas académicas y desglose visual.\n"
        "2. Comunicación funcional y autodefensa (FCT) mediante tarjetas de escritorio.\n"
        "3. Soporte de regulación sensorial con auriculares de cancelación de ruido.\n"
        "4. Sistema de reforzamiento positivo multi-nivel (PBIS)."
    )
  else:
    synth_en = (
        "Based on triangulated assessment data (65% ABC Data + 25% DSP/Vocational"
        " Interviews + 10% QABF Assessment), the adult client (Ages 21+) exhibits"
        " behaviors maintained by vocational task escape, tangible access disputes"
        " over shared community media, and elopement during unexpected alarm"
        " shifts.\n\n"
        "Comprehensive Personalized Clinical Recommendations:\n"
        "1. Person-Centered Vocational Adaptation: Provide structured visual"
        " schedules, adjustable assembly quotas, and predictable transition"
        " warnings within adult day and workshop settings.\n"
        "2. Adult Self-Advocacy & Conflict Negotiation: Train the client in"
        " adult communication scripts and choice-board negotiation protocols"
        " to manage shared community spaces and peer disputes respectfully.\n"
        "3. Environmental Safety & Elopement Prevention: Establish clear"
        " environmental safeguards, visual alarm preparation protocols, and"
        " rapid staff support procedures during schedule shifts.\n"
        "4. Positive Adult Reinforcement: Reinforce vocational endurance and"
        " appropriate self-advocacy using person-centered reinforcement."
    )
    synth_zh = (
        "基于三方交叉验证评估数据（65% ABC 数据 + 25% DSP/职业访谈 + 10% QABF 量表），该成年客户（21岁以上）"
        "表现出逃避职业组装任务、围绕共享社区媒体的物品争夺，以及在突发警报调整时擅离机构的行为。\n\n"
        "个性化综合临床建议：\n"
        "1. 以人为本的职业调整：在成人日间及工场环境中提供结构化视觉日程、可调节的工作配额以及可预测的转换预警。\n"
        "2. 成人自我倡导与冲突协商：培训客户使用成人沟通话术和选择板协商协议，以尊重的方式管理共享社区空间与同伴争议。\n"
        "3. 环境安全与擅离预防：建立清晰的环境安全保障、视觉警报准备协议以及日程调整期间的快速员工支持程序。\n"
        "4. 成人正向强化：利用以人为本的强化策略，巩固职业耐受力与恰当的自我倡导行为。"
    )
    synth_es = (
        "Basado en datos triangulados (65% ABC + 25% Entrevistas DSP + 10% QABF),"
        " el cliente adulto (21+ años) muestra conductas de escape vocacional,"
        " disputas de tangibles y fugas ante alarmas.\n\n"
        "Recomendaciones clínicas personalizadas:\n"
        "1. Adaptación vocacional centrada en la persona y horarios visuales.\n"
        "2. Autodefensa adulta y negociación de conflictos mediante tableros de opciones.\n"
        "3. Seguridad ambiental y prevención de fugas ante cambios de alarma.\n"
        "4. Reforzamiento positivo enfocado en la autonomía adulta."
    )

  add_bi_item(
      doc,
      "Clinical Synthesis & Next Steps",
      synth_en,
      (
          "临床综合结论与后续步骤"
          if lang_mode == "zh"
          else "Síntesis clínica y próximos pasos"
          if lang_mode == "es"
          else None
      ),
      (
          synth_zh
          if lang_mode == "zh"
          else synth_es
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  # Section 6: Clinical Responsibility Notice (Appended at the end)
  add_bi_heading(
      doc,
      1,
      "6. Clinical Responsibility Notice",
      (
          "6. 临床责任声明"
          if lang_mode == "zh"
          else "6. Aviso de responsabilidad clínica"
          if lang_mode == "es"
          else None
      ),
  )
  add_bi_item(
      doc,
      "Tool Positioning & Clinical Responsibility",
      (
          "This report is generated by an automated tool serving strictly as a"
          " First-Draft Synthesizer. The ultimate clinical judgment, validation"
          " of data, and formal signature responsibility (Clinical"
          " Responsibility) rest entirely with the qualified Board Certified"
          " Behavior Analyst (BCBA) or Licensed Behavior Analyst (LBA)"
          " reviewing and finalizing this document."
      ),
      (
          "工具定位与临床责任"
          if lang_mode == "zh"
          else "Posicionamiento y responsabilidad"
          if lang_mode == "es"
          else None
      ),
      (
          "本报告由自动化工具生成，仅作为初稿合成器 (First-Draft Synthesizer)。"
          "最终的临床判断、数据验证以及正式签字确认的临床责任 (Clinical Responsibility)"
          " 完全由审核并定稿该文件的合格行为分析师 (BCBA 或 LBA) 承担。"
          if lang_mode == "zh"
          else "Este informe es un sintetizador de primer borrador. La responsabilidad clínica final recae en el BCBA o LBA."
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  append_fba_evidence_references(doc, lang_mode)

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 8. Comprehensive Enriched BIP Generator
# ==========================================
def generate_exact_bip_doc(cohort_key, lang_choice, template_choice="School-Based IEP"):
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
    r_tr = p_tr.add_run("[行为干预计划 (BIP) Comprehensive Draft]")
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)
  elif lang_mode == "es":
    p_tr = doc.add_paragraph()
    r_tr = p_tr.add_run(
        "[Plan de Intervención Conductual (BIP) Borrador Completo]"
    )
    r_tr.italic = True
    r_tr.font.color.rgb = RGBColor(100, 100, 100)

  # Section 1
  add_bi_heading(
      doc,
      1,
      "1. Student Info & Administrative Summary",
      (
          "1. 学生/客户信息与行政摘要"
          if lang_mode == "zh"
          else "1. Información del estudiante y resumen"
          if lang_mode == "es"
          else None
      ),
  )
  build_compact_demographics_table(doc, c_meta, lang_mode, template_choice)
  is_school_bip = "School" in template_choice
  add_bi_item(
      doc,
      "Plan Orientation & Service Focus",
      (
          "Organized around School / Classroom & IEP Alignment, prioritizing"
          " classroom routines, transitions, teacher/paraeducator prompts, and"
          " IEP objectives."
          if is_school_bip
          else "Organized around Clinical / Home & Community Supports,"
          " prioritizing caregiver routines, home/community settings, and"
          " structured teaching / reinforcement contingencies."
      ),
      (
          "计划定位与服务侧重"
          if lang_mode == "zh"
          else "Orientación del plan y enfoque"
          if lang_mode == "es"
          else None
      ),
      (
          (
              "围绕 School / Classroom & IEP Alignment 组织，侧重"
              " classroom routines, transitions, teacher/paraeducator prompts, IEP"
              " objectives。"
              if is_school_bip
              else "围绕 Clinical / Home & Community Supports 组织，侧重"
              " caregiver routines, home/community settings, structured teaching /"
              " reinforcement contingencies。"
          )
          if lang_mode == "zh"
          else (
              "Organizado en torno a School / Classroom & IEP Alignment."
              if is_school_bip
              else "Organizado en torno a Clinical / Home & Community Supports."
          )
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  # Section 2: Behavior Functions & FERB Breakdown
  add_bi_heading(
      doc,
      1,
      "2. Target Behaviors, Functions & Replacement Skills (FERB)",
      (
          "2. 目标行为、行为功能与替代技能 (FERB) 逐项拆解"
          if lang_mode == "zh"
          else "2. Conductas objetivo, funciones y conductas de reemplazo (FERB)"
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
        "Validated Function",
        b["hypothesis"],
        (
            "确证行为功能"
            if lang_mode == "zh"
            else "Función validada"
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
    add_bi_item(
        doc,
        "Functionally Equivalent Replacement Behavior (FERB)",
        b["ferb"],
        (
            "功能性替代行为 (FERB)"
            if lang_mode == "zh"
            else "Conducta de reemplazo funcionalmente equivalente (FERB)"
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

  # Section 3: Proactive / Antecedent Strategies
  add_bi_heading(
      doc,
      1,
      "3. Proactive & Antecedent Modifications (Prevention)",
      (
          "3. 前因调整与预防策略 (统一整合)"
          if lang_mode == "zh"
          else "3. Modificaciones proactivas y antecedentes"
          if lang_mode == "es"
          else None
      ),
  )
  add_bi_item(
      doc,
      "3.1 Environmental & Priming Adaptations",
      (
          (
              "• Provide 2-minute and 1-minute transition warnings across"
              " classroom routines and transitions using teacher/paraeducator"
              " prompts aligned to IEP objectives.\n• Offer noise-canceling"
              " headphones or move student to low-stimulation sensory area prior"
              " to task demands.\n• Break academic/vocational tasks into small,"
              " visual chunks (2-3 items per strip)."
          )
          if is_school_bip
          else (
              "• Provide 2-minute and 1-minute transition warnings across"
              " caregiver routines and home/community settings using structured"
              " teaching / reinforcement contingencies.\n• Offer noise-canceling"
              " headphones or move client to low-stimulation sensory area prior"
              " to task demands.\n• Break academic/vocational tasks into small,"
              " visual chunks (2-3 items per strip)."
          )
      ),
      (
          "3.1 环境调整与预先提示"
          if lang_mode == "zh"
          else "3.1 Adaptaciones ambientales"
          if lang_mode == "es"
          else None
      ),
      (
          (
              (
                  "• 在 classroom routines 与 transitions 前提供 2分钟 及 1分钟"
                  " 提示，结合 teacher/paraeducator prompts 并对齐 IEP"
                  " objectives。\n• 在任务开始前主动提供降噪耳机或移至低刺激感官区域。\n•"
                  " 将学业/职业任务拆解为小步子视觉单元（每条2-3个小任务）。"
              )
              if is_school_bip
              else (
                  "• 在 caregiver routines 与 home/community settings"
                  " 前提供 2分钟 及 1分钟 提示，结合 structured teaching /"
                  " reinforcement contingencies。\n•"
                  " 在任务开始前主动提供降噪耳机或移至低刺激感官区域。\n•"
                  " 将学业/职业任务拆解为小步子视觉单元（每条2-3个小任务）。"
              )
          )
          if lang_mode == "zh"
          else (
              (
                  "• Proporcionar avisos de 2 min y 1 min en classroom"
                  " routines y transitions con teacher/paraeducator prompts y"
                  " IEP objectives.\n• Ofrecer auriculares de cancelación de"
                  " ruido antes de tareas.\n• Dividir tareas en pequeños"
                  " bloques visuales."
              )
              if is_school_bip
              else (
                  "• Proporcionar avisos de 2 min y 1 min en caregiver"
                  " routines y home/community settings con structured"
                  " teaching / reinforcement contingencies.\n• Ofrecer"
                  " auriculares de cancelación de ruido antes de tareas.\n•"
                  " Dividir tareas en pequeños bloques visuales."
              )
          )
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  # Section 4: Replacement Behaviors Protocols
  add_bi_heading(
      doc,
      1,
      "4. Replacement Behaviors & Functional Communication Training (FCT)",
      (
          "4. 替代行为与功能性沟通训练 (FCT)"
          if lang_mode == "zh"
          else "4. Entrenamiento en Comunicación Funcional (FCT)"
          if lang_mode == "es"
          else None
      ),
  )
  add_bi_item(
      doc,
      "4.1 Functional Communication Protocols",
      (
          "• Primary FCT Skill: Prompt client to press an AAC button or hand a"
          " PECS icon for 'Break' or 'Help' upon initial sign of"
          " frustration.\n• Systematic Prompting: Use Most-to-Least physical"
          " assistance, fading rapidly to gestural/visual prompts within 10"
          " days."
      ),
      (
          "4.1 功能性沟通协议"
          if lang_mode == "zh"
          else "4.1 Protocolos de comunicación funcional"
          if lang_mode == "es"
          else None
      ),
      (
          "• 核心 FCT 技能：教导客户在产生烦躁情绪萌芽时，按下 AAC 沟通按键或递交 PECS"
          " 卡片表达“休息”或“帮助”。\n• 辅助渐退策略：使用由多到少（Most-to-Least）物理辅助，并在"
          " 10 天内快速渐退至手势或视觉提示。"
          if lang_mode == "zh"
          else "• Habilidad FCT principal: Indicar al cliente que presione el"
          " botón AAC o entregue la tarjeta PECS 'Descanso'.\n• Fading de"
          " ayudas: De mayor a menor asistencia en 10 días."
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  # Section 5: Reinforcement Strategies
  add_bi_heading(
      doc,
      1,
      "5. Reinforcement Protocols",
      (
          "5. 强化策略协议"
          if lang_mode == "zh"
          else "5. Protocolos de Reforzamiento"
          if lang_mode == "es"
          else None
      ),
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
      (
          "5.1 替代行为区别性强化 (DRA)"
          if lang_mode == "zh"
          else "5.1 Reforzamiento Diferencial de Conducta Alternativa (DRA)"
          if lang_mode == "es"
          else None
      ),
      (
          "• 在习得阶段，只要客户按下 AAC 表达“休息”，必须在 3 秒内 100%"
          " 满足其休息请求。\n•"
          " 将逃避任务与高度热情的情感口头表扬结合（如：“太棒了，你自己按按键说要休息！”）。"
          if lang_mode == "zh"
          else "• Cumplimiento inmediato (en 3 segundos) del 100% con la"
          " solicitud de 'Descanso'.\n• Acompañar el descanso con el elogio"
          " verbal entusiasta."
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  # Section 6: Response Strategies
  add_bi_heading(
      doc,
      1,
      "6. Reactive Response Protocols & Extinction",
      (
          "6. 目标行为回应与消退策略"
          if lang_mode == "zh"
          else "6. Protocolos de respuesta reactiva y extinción"
          if lang_mode == "es"
          else None
      ),
  )
  add_bi_item(
      doc,
      "6.1 Extinction & Neutral Physical Blocking",
      (
          "• Escape Extinction: Maintain neutral expression, avoid eye contact,"
          " minimize verbal dialogue during problem behavior.\n• Physical"
          " Blocking: Promptly and softly block any SIB, slapping, or"
          " floor-dropping using foam pads to prevent injury without giving"
          " emotional feedback."
      ),
      (
          "6.1 消退与中立物理阻挡"
          if lang_mode == "zh"
          else "6.1 Extinción y bloqueo físico neutral"
          if lang_mode == "es"
          else None
      ),
      (
          "•"
          " 逃避消退：在问题行为发生时，保持平静中立表情，避免眼神接触，不进行长篇大论训诫。\n•"
          " 物理阻挡：若出现自伤或攻击行为，使用软垫迅速柔和阻挡，确保安全的同时不给予额外的言语或情感反馈。"
          if lang_mode == "zh"
          else "• Extinción de escape: Mantener expresión neutral, evitar contacto"
          " visual.\\n• Bloqueo físico: Bloquear SIB o bofetadas de forma suave"
          " usando colchonetas."
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  # Section 7: Safety & Data (Updated with Crisis Safety Content)
  add_bi_heading(
      doc,
      1,
      "7. Crisis Safety & Treatment Fidelity",
      (
          "7. 危机安全预案与执行忠实度"
          if lang_mode == "zh"
          else "7. Seguridad en crisis y fidelidad del tratamiento"
          if lang_mode == "es"
          else None
      ),
  )
  add_bi_item(
      doc,
      "Crisis Safety Protocols & Emergency Procedures",
      (
          "• Crisis Plan Activation: Implemented strictly when behaviors escalate to severe self-injury, physical aggression, or elopement into unsafe zones where de-escalation and FCT are temporarily insufficient.\n"
          "• Environmental Safety & Clearing: Immediately clear surrounding hazardous objects, clear peers/bystanders from the immediate radius, and secure exits if elopement risk is present.\n"
          "• De-escalation & Restraint Policy: Utilize least-restrictive crisis management techniques. Physical holding or restrictive interventions are reserved exclusively as a last resort for imminent, severe physical harm when all proactive and alternative interventions have failed, and must be documented immediately per facility policy."
      ),
      (
          "危机安全预案与紧急处置程序"
          if lang_mode == "zh"
          else "Protocolos de seguridad en crisis y procedimientos de emergencia"
          if lang_mode == "es"
          else None
      ),
      (
          "• 危机预案启动：当行为升级为严重自伤、身体攻击或冲入危险区域，且前因调整与 FCT 暂时无法缓解时严格启动。\n"
          "• 环境安全与清场：立即清理周边的危险物品，将同伴/旁观者疏散至安全半径之外，若存在擅离风险则锁定出口。\n"
          "• 降温去激化与约束政策：采用限制性最小的危机管理技术。身体约束或限制性干预仅在面临迫在眉睫的严重身体伤害、且所有前因及替代干预均告失败时作为最后手段使用，并须根据机构规定立即记录。"
          if lang_mode == "zh"
          else "• Activación del plan de crisis: Se implementa estrictamente cuando las conductas escalan a autolesión severa, agresión física o fuga a zonas inseguras.\n"
          "• Seguridad ambiental y despeje: Retirar objetos peligrosos cercanos y despejar a los espectadores.\n"
          "• Desescalada y política de contención: Utilizar técnicas de manejo de crisis menos restrictivas. La contención física se reserva exclusivamente como último recurso ante daño físico inminente."
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )
  add_bi_item(
      doc,
      "Data Collection & Fidelity Protocols",
      (
          "• RBTs will record daily frequency/duration of target behaviors and"
          " independent FCT requests.\n• BCBA will conduct weekly treatment"
          " fidelity observations using a 10-point checklist."
      ),
      (
          "数据收集与忠实度核查"
          if lang_mode == "zh"
          else "Recolección de datos y fidelidad"
          if lang_mode == "es"
          else None
      ),
      (
          "• RBT 每日记录目标行为的发生频率/持续时间及 FCT 独立使用次数。\n• BCBA"
          " 每周使用 10 项标准核查表进行 1:1 干预忠实度评估。"
          if lang_mode == "zh"
          else "• Los RBTs registrarán la frecuencia/duración diaria.\\n• El"
          " BCBA realizará observaciones semanales de fidelidad."
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  # Section 8: Staff Training and Monitoring
  add_bi_heading(
      doc,
      1,
      "8. Staff Training and Monitoring",
      (
          "8. 人员培训、执行监督与计划复审机制"
          if lang_mode == "zh"
          else "8. Capacitación del personal y supervisión"
          if lang_mode == "es"
          else None
      ),
  )

  st_en = (
      "• Responsible Party: Board Certified Behavior Analyst (BCBA) / Licensed"
      " Behavior Analyst (LBA).\n• Training Process: Initial 2-hour didactic"
      " module followed by 1:1 behavioral skills training (BST: Instructions,"
      " Modeling, Rehearsal, and Feedback) with all direct intervention staff"
      " (RBTs/Therapists).\n• Observation & Fidelity Frequency: Direct"
      " weekly observations utilizing a 10-point Treatment Fidelity Checklist"
      " to ensure implementation accuracy (>90% compliance required).\n• Plan"
      " Review Schedule: Formal clinical review scheduled every 90 days, or"
      " immediately if target behaviors show a 20% spike in frequency."
  )
  st_zh = (
      "• 责任人：管辖 BCBA 行为分析师 / LBA 执照行为分析师。\n• 培训流程：实施 2"
      " 小时初始理论培训，随后进行 1:1 行为技能训练 (BST"
      " 模式：讲解、示范、演练与即时反馈)，覆盖所有直属 RBT / 干预人员。\n•"
      " 观察与忠实度核查频率：每周进行一次 1:1 现场观察，使用 10"
      " 项行为忠实度清单核查（要求达标率 >90%）。\n• 计划复审时间表：每 90"
      " 天进行一次正式临床复审，若目标行为频率发生 >20% 的异常激增则触发即时复审。"
  )
  st_es = (
      "• Persona responsable: BCBA / LBA.\n• Proceso de capacitación: Módulo"
      " de 2 horas seguido de Entrenamiento en Habilidades Conductuales (BST:"
      " Instrucciones, Modelado, Ensayo y Retroalimentación).\n• Frecuencia de"
      " observación: Observaciones semanales con lista de cotejo de fidelidad"
      " del 10-pt (requerido >90%).\n• Programación de revisión: Revisión"
      " clínica cada 90 días."
  )

  add_bi_item(
      doc,
      "Staff Training & Fidelity Monitoring Plan",
      st_en,
      (
          "人员培训与忠实度监控计划"
          if lang_mode == "zh"
          else "Plan de capacitación y monitoreo"
          if lang_mode == "es"
          else None
      ),
      (
          st_zh
          if lang_mode == "zh"
          else st_es
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )
  if is_school_bip:
    add_bi_item(
        doc,
        "Classroom & IEP Goal Alignment / Staff & Paraeducator Implementation",
        (
            "• Align target behaviors, FCT, and reinforcement with IEP"
            " objectives across classroom routines and transitions.\n• Train"
            " teachers and paraeducators on consistent teacher/paraeducator"
            " prompts, seating/desk placement, and pre-transition warnings."
        ),
        "课堂与 IEP 目标对齐 / 教师与助教实施" if lang_mode == "zh" else "Alineación IEP e implementación del personal" if lang_mode == "es" else None,
        (
            "• 将目标行为、FCT 与强化与 IEP objectives 对齐，贯穿 classroom"
            " routines 和 transitions。\n• 培训教师与助教规范使用"
            " teacher/paraeducator prompts、课桌安排与转换前预警。"
            if lang_mode == "zh"
            else "• Alinear con IEP objectives en classroom routines y"
            " transitions.\n• Capacitar a maestros y paraeducadores en"
            " teacher/paraeducator prompts."
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )
  else:
    add_bi_item(
        doc,
        "Caregiver Coaching & Home/Community Generalization",
        (
            "• Coach caregivers on caregiver routines, structured teaching /"
            " reinforcement contingencies, and consistent FCT responses.\n•"
            " Program generalization across home/community settings and daily"
            " routines."
        ),
        "照护者实操指导 / 家庭与社区泛化" if lang_mode == "zh" else "Coaching a cuidadores y generalización" if lang_mode == "es" else None,
        (
            "• 指导照护者落实 caregiver routines、structured teaching /"
            " reinforcement contingencies 与一致的 FCT 回应。\n• 在"
            " home/community settings 与日常作息中推进泛化。"
            if lang_mode == "zh"
            else "• Capacitar cuidadores en caregiver routines y structured"
            " teaching / reinforcement contingencies.\n• Generalizar en"
            " home/community settings."
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )

  # Section 9: Clinical Responsibility Notice (Appended at the end)
  add_bi_heading(
      doc,
      1,
      "9. Clinical Responsibility Notice",
      (
          "9. 临床责任声明"
          if lang_mode == "zh"
          else "9. Aviso de responsabilidad clínica"
          if lang_mode == "es"
          else None
      ),
  )
  add_bi_item(
      doc,
      "Tool Positioning & Clinical Responsibility",
      (
          "This Behavior Intervention Plan is generated by an automated tool"
          " serving strictly as a First-Draft Synthesizer. The ultimate"
          " clinical judgment, validation of intervention strategies, and"
          " formal signature responsibility (Clinical Responsibility) rest"
          " entirely with the qualified Board Certified Behavior Analyst (BCBA)"
          " or Licensed Behavior Analyst (LBA)."
      ),
      (
          "工具定位与临床责任"
          if lang_mode == "zh"
          else "Posicionamiento y responsabilidad"
          if lang_mode == "es"
          else None
      ),
      (
          "本行为干预计划由自动化工具生成，仅作为初稿合成器 (First-Draft Synthesizer)。"
          "最终的临床判断、干预策略验证以及正式签字确认的临床责任 (Clinical Responsibility)"
          " 完全由审核并定稿该文件的合格行为分析师 (BCBA 或 LBA) 承担。"
          if lang_mode == "zh"
          else "Este plan es un sintetizador de primer borrador. La responsabilidad clínica final recae en el BCBA o LBA."
          if lang_mode == "es"
          else None
      ),
      lang_mode,
  )

  append_bip_evidence_references(doc, lang_mode)

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio


# ==========================================
# 9. Action Buttons & Template Selection
# ==========================================
st.markdown("### 3️⃣ Template, Target Language & Formulate / Download Actions")

if (
    uploaded_abc is None
    or uploaded_interview is None
    or uploaded_qabf is None
):
  st.markdown(
      """
      <div style="background-color: #FEF5E7; border: 2px solid #D68910; border-left: 8px solid #D68910; padding: 0.75rem 1.0rem; border-radius: 6px; margin-bottom: 0.85rem; color: #7E5109; font-size: 0.95rem; font-weight: 600; line-height: 1.45;">
          Sample Data in Use: Not all three assessment files (ABC, Interview, QABF) have been uploaded. This draft is currently generated using built-in sample data. Please upload all required assessment files to generate a case-specific FBA/BIP.
      </div>
      """,
      unsafe_allow_html=True,
  )

col_tmpl, col_lang, col_action1, col_action2 = st.columns([1.2, 1.2, 1.3, 1.3])

with col_tmpl:
  template_choice = st.radio(
      "Select Report Template:",
      options=[
          "School-Based IEP",
          "Commercial / Clinical",
      ],
      index=0,
  )

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
    selected_cohort_key, report_lang, active_behaviors, template_choice
)
bip_docx_bytes = generate_exact_bip_doc(
    selected_cohort_key, report_lang, template_choice
)

tmpl_tag = "School_IEP" if "School" in template_choice else "Commercial_Clinical"

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
      label="⚡ Formulate & Download De-Identified FBA Draft (.docx)",
      data=fba_docx_bytes,
      file_name=(
          f"DeIdentified_FBA_Draft_{current_meta['file_tag']}_{tmpl_tag}_{lang_tag}.docx"
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
          f"DeIdentified_BIP_Draft_{current_meta['file_tag']}_{tmpl_tag}_{lang_tag}.docx"
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
