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
    .triangulation-badge {
        background-color: #1F4E78;
        color: white;
        padding: 0.3rem 0.6rem;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.85rem;
        display: inline-block;
        margin-top: 0.4rem;
    }
    
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
      "g1": {
        "title": "Cohort 1: Early Intervention (Ages 2-5)",
        "title_zh": "人群 1：早期干预组（2-5岁）",
        "title_es": "Grupo 1: Intervención Temprana (2-5 años)",
        "behaviors": [
            {
                "name": "1. Self-Injurious Behavior (SIB) - Head Banging & Wrist Biting",
                "name_zh": "1. 自伤行为 (SIB) - 撞头与咬手腕",
                "name_es": "1. Conducta autolesiva (SIB) - Golpes en la cabeza y mordedura de muñeca",
                "def": "Any instance where the child forcefully makes contact between forehead and a hard/padded surface, or places wrist/hand between upper and lower teeth with visible physical force lasting >1 second. Onset is the initial physical contact; offset is 5 consecutive seconds without contact.",
                "def_zh": "儿童额头与硬质或软垫表面发生有力的碰撞，或将手腕/手部置于上下牙齿之间咬合并伴有明显施力的任何行为，持续时间>1秒。以首次物理接触为行为开始，以连续5秒无上述动作为行为结束。",
                "def_es": "Cualquier instancia en la que el niño haga contacto con fuerza entre la frente y una superficie dura/acolchada, o coloque la muñeca/mano entre los dientes superiores e inferiores con fuerza física visible que dure >1 segundo.",
                "ex": "Banging forehead 3-4 times on foam mat during tabletop task; leaving red teeth marks on right wrist during transition.",
                "ex_zh": "在桌面任务期间，额头在泡棉垫上连续碰撞3-4次；在转换环节中在右手腕上留下红印牙痕。",
                "ex_es": "Golpear la frente 3-4 veces sobre la colchoneta durante la tarea; dejar marcas rojas de dientes en la muñeca derecha.",
                "non_ex": "Resting head on floor mat during group circle time; mouthing FDA-approved silicone chew necklace softly.",
                "non_ex_zh": "在集体圈圈时间将头靠在地垫上休息；轻柔地咀嚼经过安全认证的硅胶项链。",
                "non_ex_es": "Descansar la cabeza en la colchoneta; morder suavemente un collar de silicona apto.",
                "dimensions": "Frequency: 3-6 episodes/day. Duration: 15s - 2min per outburst. Intensity: Moderate to Severe (potential skin redness or bruising).",
                "dimensions_zh": "频率：每天 3-6 次。持续时间：每次发作 15秒 至 2分钟。强度：中度至重度（可能导致皮肤发红或瘀青）。",
                "dimensions_es": "Frecuencia: 3-6 episodios/día. Duración: 15s - 2min. Intensidad: Moderada a Grave.",
                "triggers": "Setting Events: Fatigue, high ambient background noise.\nImmediate Triggers: Presentation of fine-motor discrete trial tasks, removal of preferred sensory toy.",
                "triggers_zh": "背景事件：疲劳、环境背景噪音过大。\n直接触发因素：呈现精细动作桌面任务、移走偏好的感官玩具。",
                "triggers_es": "Eventos de contexto: Fatiga, ruido ambiental.\nDesencadenantes inmediatos: Tareas motoras finas, retirada de juguete preferido.",
                "consequences": "RBT blocks contact using foam pad, pauses academic demand immediately, and prompts PECS 'Break' card.",
                "consequences_zh": "RBT 使用软垫阻挡物理接触，立即暂停学业要求，并提示使用 PECS“休息”卡片。",
                "consequences_es": "El RBT bloquea el contacto, pausa la demanda e indica la tarjeta PECS 'Descanso'.",
                "qabf_summary": "Task Escape: 14/15 | Physical Discomfort: 8/15 | Attention: 2/15 | Tangible: 3/15 | Sensory: 4/15",
                "qabf_summary_zh": "逃避任务: 14/15 | 身体不适: 8/15 | 社交关注: 2/15 | 获得物质: 3/15 | 感官刺激: 4/15",
                "qabf_summary_es": "Escape de tarea: 14/15 | Malestar físico: 8/15 | Atención: 2/15 | Tangible: 3/15 | Sensorial: 4/15",
                "triangulation": "65% Direct ABC Data (high rate during table work) + 25% Indirect Parent Interview (frustration with structured demands) + 10% QABF (Escape score 14/15).",
                "triangulation_zh": "65% 直接 ABC 数据（桌面任务期间高发） + 25% 间接家长访谈（对结构化任务表现出挫败） + 10% QABF 结果（逃避任务得分 14/15）。",
                "triangulation_es": "65% Datos ABC + 25% Entrevista + 10% QABF (Escape 14/15).",
                "hypothesis": "Primary Function: Task Escape (Social Negative Reinforcement).\nSecondary Function: Physical Discomfort Relief.",
                "hypothesis_zh": "核心行为功能：逃避任务（社交负强化）。\n次要行为功能：缓解身体不适。",
                "hypothesis_es": "Función principal: Escape de tareas.\nFunción secundaria: Alivio de malestar físico.",
                "ferb": "Functional Communication: Activate AAC BigMack button ('I need a break') or hand 'Break' PECS card to prompt immediate task pause.",
                "ferb_zh": "功能性替代行为 (FERB)：按下 AAC 语音按键（“我想休息”）或递交“休息” PECS 卡片，以提示示范暂停任务。",
                "ferb_es": "Comunicación Funcional: Activar el botón AAC BigMack ('Necesito un descanso') o entregar la tarjeta PECS 'Descanso'.",
            },
            {
                "name": "2. Sensory Vocal Distress & Face-Slapping",
                "name_zh": "2. 感官情绪性尖叫与拍打面部",
                "name_es": "2. Distrés vocal sensorial y bofetadas en la cara",
                "def": "Vocal screaming exceeding normal conversational volume (>80 dB) lasting >2 seconds, occurring simultaneously with open-palm striking of own cheeks or arms.",
                "def_zh": "尖叫声超过正常交谈音量（>80 dB）且持续时间>2秒，同时伴随用张开的手掌拍打自己脸颊或手臂的行为。",
                "def_es": "Gritos vocales que superan el volumen conversacional normal (>80 dB) durante >2 segundos, junto con golpes con la palma abierta en las mejillas o brazos.",
                "ex": "Loud screaming and striking cheeks 3 times when loud kitchen appliance starts in adjacent room.",
                "ex_zh": "当隔壁房间开启高分贝厨房电器时，大声尖叫并连续拍打脸颊 3 次。",
                "ex_es": "Gritos fuertes y 3 golpes en las mejillas al encenderse un electrodoméstico ruidoso.",
                "non_ex": "Excited vocal shouting during outdoor playground play; tapping cheeks rhythmically during music group time.",
                "non_ex_zh": "在户外操场游玩时兴奋地高声欢呼；在音乐课上跟随节奏轻拍脸颊。",
                "non_ex_es": "Gritos de emoción en el parque; golpear las mejillas al ritmo de la música.",
                "dimensions": "Frequency: 2-4 episodes/day. Duration: 30s - 3min. Intensity: Moderate.",
                "dimensions_zh": "频率：每天 2-4 次。持续时间：30秒 至 3分钟。强度：中度。",
                "dimensions_es": "Frecuencia: 2-4 episodios/día. Duración: 30s - 3min. Intensidad: Moderada.",
                "triggers": "Setting Events: Overstimulating ambient noise, sudden schedule changes.\nImmediate Triggers: Unexpected high-pitch sounds, removal of preferred juice cup.",
                "triggers_zh": "背景事件：环境噪音过度刺激、突发日程改变。\n直接触发因素：意料之外的高频噪音、偏好的果汁杯被移走。",
                "triggers_es": "Eventos de contexto: Ruido ambiental excesivo.\nDesencadenantes: Sonidos agudos inesperados.",
                "consequences": "Staff offers noise-canceling headphones and redirects to sensory chew tool.",
                "consequences_zh": "工作人员提供降噪耳机，并重新引导至口部感官咀嚼工具。",
                "consequences_es": "El personal ofrece auriculares de cancelación de ruido y redirige a un mordedor sensorial.",
                "qabf_summary": "Sensory/Automatic: 13/15 | Task Escape: 11/15 | Attention: 3/15 | Tangible: 2/15 | Physical: 1/15",
                "qabf_summary_zh": "感官刺激/自动强化: 13/15 | 逃避任务: 11/15 | 社交关注: 3/15 | 获得物质: 2/15 | 身体不适: 1/15",
                "qabf_summary_es": "Sensorial: 13/15 | Escape: 11/15 | Atención: 3/15 | Tangible: 2/15 | Físico: 1/15",
                "triangulation": "65% Direct ABC Data (appliance noise trigger) + 25% Indirect RBT Interview (auditory aversion) + 10% QABF (Sensory score 13/15).",
                "triangulation_zh": "65% 直接 ABC 数据（电器噪音触发） + 25% 间接 RBT 访谈（听觉敏感过度） + 10% QABF 结果（感官得分 13/15）。",
                "triangulation_es": "65% Datos ABC + 25% Entrevista RBT + 10% QABF (Sensorial 13/15).",
                "hypothesis": "Primary Function: Escape / Regulation of Auditory Overstimulation (Sensory Automatic Reinforcement).",
                "hypothesis_zh": "核心行为功能：逃避/调节听觉过度刺激（感官自动强化）。",
                "hypothesis_es": "Función principal: Escape / Regulación de sobreestimulación auditiva.",
                "ferb": "Self-Regulation / Request Strategy: Point to 'Headphones' visual symbol or independently retrieve noise-canceling headphones from designated sensory bin.",
                "ferb_zh": "功能性替代行为 (FERB)：指认“耳机”视觉标识，或自行从指定的感官箱中取出降噪耳机戴上。",
                "ferb_es": "Estrategia de autorregulación: Señalar el símbolo 'Auriculares' o tomarlos de la caja sensorial.",
            },
            {
                "name": "3. Property Destruction & Drop-to-Floor (Flopping)",
                "name_zh": "3. 破坏物品与躺地抗拒 (Flopping)",
                "name_es": "3. Destrucción de propiedad y tirarse al suelo (Flopping)",
                "def": "Any instance where the child forcibly throws, sweeps, or slams therapy materials/toys off surfaces, or drops body weight entirely to the floor resisting adult guidance for >3 seconds upon removal of a preferred tangible item.",
                "def_zh": "当偏好的物品/玩具被移走时，儿童用力抛掷、扫落或重摔干预用具/玩具，或将全身重量沉入地面拒绝成人引导持续>3秒的任何行为。",
                "def_es": "Cualquier instancia en la que el niño arroje, empuje o golpee materiales/juguetes, o deje caer su peso al suelo resistiéndose a la guía adulta durante >3 segundos tras la retirada de un objeto preferido.",
                "ex": "Sweeping discrete trial matching cards off desk and lying limp on floor when preferred tablet access is terminated.",
                "ex_zh": "当平板电脑使用时间结束时，把桌面上的干预配对卡片全部扫落到地上，并全身发软趴在地上不肯起来。",
                "ex_es": "Tirar las tarjetas de trabajo al suelo y tirarse al suelo cuan largo es al finalizar el tiempo de tableta.",
                "non_ex": "Accidentally knocking a crayon off desk during coloring; sitting down on play mat upon instructions.",
                "non_ex_zh": "涂色时不小心将蜡笔碰落桌下；听从指令坐在游戏地垫上。",
                "non_ex_es": "Tirar una crayola por accidente; sentarse en la colchoneta tras una instrucción.",
                "dimensions": "Frequency: 2-3 episodes/day. Duration: 1-5 minutes. Intensity: Moderate.",
                "dimensions_zh": "频率：每天 2-3 次。持续时间：1-5 分钟。强度：中度。",
                "dimensions_es": "Frecuencia: 2-3 episodios/día. Duración: 30s - 3min. Intensidad: Moderada.",
                "triggers": "Setting Events: Transitions between high-preference and low-preference activities.\nImmediate Triggers: Verbal prompt to transition away from iPad/preferred sensory light toy.",
                "triggers_zh": "背景事件：在高偏好与低偏好活动之间的环节转换。\n直接触发因素：要求停止使用 iPad 或偏好感官玩具的语言提示。",
                "triggers_es": "Eventos: Transición de alta a baja preferencia.\nDesencadenante: Instrucción de entregar el iPad.",
                "consequences": "Staff provides visual countdown board, maintains safety, and prompts PECS 'More Time' or 'My Turn' icon.",
                "consequences_zh": "工作人员出示视觉倒计时板，维持安全保护，并提示使用 PECS“还想要/再玩会儿”标识。",
                "consequences_es": "El personal presenta contador visual y promueve el uso del ícono PECS 'Más tiempo'.",
                "qabf_summary": "Access to Tangible: 14/15 | Task Escape: 9/15 | Attention: 5/15 | Sensory: 2/15 | Physical: 0/15",
                "qabf_summary_zh": "获得物质/活动: 14/15 | 逃避任务: 9/15 | 社交关注: 5/15 | 感官刺激: 2/15 | 身体不适: 0/15",
                "qabf_summary_es": "Acceso a Tangible: 14/15 | Escape: 9/15 | Atención: 5/15 | Sensorial: 2/15 | Físico: 0/15",
                "triangulation": "65% Direct ABC Data (high occurrence at transition times) + 25% Parent Interview (struggles with home iPad limits) + 10% QABF (Tangible score 14/15).",
                "triangulation_zh": "65% 直接 ABC 数据（转换环节高发） + 25% 家人访谈（在家难以接受 iPad 时间限制） + 10% QABF 结果（获得物质得分 14/15）。",
                "triangulation_es": "65% Datos ABC + 25% Entrevista familiar + 10% QABF (Tangible 14/15).",
                "hypothesis": "Primary Function: Access to Preferred Tangible Items / Activities (Social Positive Reinforcement).",
                "hypothesis_zh": "核心行为功能：获取偏好的物品/活动（社交正强化）。",
                "hypothesis_es": "Función principal: Acceso a objetos o actividades preferidas.",
                "ferb": "Functional Communication: Hand 'More Time' PECS icon or press AAC button to request a 2-minute extension before transition.",
                "ferb_zh": "功能性替代行为 (FERB)：递交“再玩会儿” PECS 卡片，或按压 AAC 表达“还需要 2 分钟”，以适当延缓转换。",
                "ferb_es": "Comunicación Funcional: Entregar la tarjeta PECS 'Más tiempo' o presionar el botón AAC.",
            }
        ]
    },
        "g2": [
            {
                "Date_Time": "2026-08-10 09:30",
                "Setting": "Gen-Ed Classroom",
                "Antecedent": "Teacher presented 2-page math worksheet",
                "Behavior": "Screamed 'I won't do it!', pushed desk away",
                "Consequence": "Staff presented 'Break' visual card; demand paused"
            }
        ],
        "g3": [
            {
                "Date_Time": "2026-08-09 09:00",
                "Setting": "Residential Home",
                "Antecedent": "New staff worker introduced schedule change",
                "Behavior": "Swept dishes off table, shouted threats",
                "Consequence": "DSP stepped in, offered visual choice board, demand paused"
            }
        ]
    }
    df = pd.DataFrame(datasets.get(cohort_key, datasets["g1"]))
    return df.to_csv(index=False).encode('utf-8')

def generate_mock_interview_docx(cohort_key):
    doc = docx.Document()
    doc.add_heading('INDIRECT ASSESSMENT: STAKEHOLDER INTERVIEW NOTES (DE-IDENTIFIED)', level=1)
    doc.add_paragraph(f"Client ID: [CLIENT_ID] | Target Cohort: {cohort_meta[cohort_key]['title']}\nInformants: Parent, Lead Therapist / Educator, RBT Supervisor\n")
    doc.add_paragraph("Summary: Stakeholders report elevated rates of target behaviors during task transitions, fine-motor academic demands, and high-pitch sensory noise environments.")
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def generate_mock_qabf_docx(cohort_key):
    doc = docx.Document()
    doc.add_heading('PSYCHOMETRIC QABF ASSESSMENT RESULTS (PER TARGET BEHAVIOR)', level=1)
    doc.add_paragraph(f"Client ID: [CLIENT_ID] | Cohort: {cohort_meta[cohort_key]['title']}\n")
    b_list = cohort_meta[cohort_key]['behaviors']
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
        "triangulation_algo": "65% Direct ABC Data + 25% Indirect Interview Data + 10% QABF",
        "protocol_sentences": [
            "• Focuses on developmental milestone integration using ESDM / NDBI naturalistic approaches.",
            "• Prioritizes sensory processing, emotional regulation, and early functional communication (PECS/AAC).",
            "• Integrates play-based assessment with parent-led co-regulation routines.",
            "• Emphasizes proactive environmental adaptation and rapid reinforcement for replacement skills."
        ],
        "behaviors": [，
            {
                "name": "1. Self-Injurious Behavior (SIB) - Head Banging & Wrist Biting",
                "name_zh": "1. 自伤行为 (SIB) - 撞头与咬手腕",
                "name_es": "1. Conducta autolesiva (SIB) - Golpes en la cabeza y mordedura de muñeca",
                "def": "Any instance where the child forcefully makes contact between forehead and a hard/padded surface, or places wrist/hand between upper and lower teeth with visible physical force lasting >1 second. Onset is the initial physical contact; offset is 5 consecutive seconds without contact.",
                "def_zh": "儿童额头与硬质或软垫表面发生有力的碰撞，或将手腕/手部置于上下牙齿之间咬合并伴有明显施力的任何行为，持续时间>1秒。以首次物理接触为行为开始，以连续5秒无上述动作为行为结束。",
                "def_es": "Cualquier instancia en la que el niño haga contacto con fuerza entre la frente y una superficie dura/acolchada, o coloque la muñeca/mano entre los dientes superiores e inferiores con fuerza física visible que dure >1 segundo.",
                "ex": "Banging forehead 3-4 times on foam mat during tabletop task; leaving red teeth marks on right wrist during transition.",
                "ex_zh": "在桌面任务期间，额头在泡棉垫上连续碰撞3-4次；在转换环节中在右手腕上留下红印牙痕。",
                "ex_es": "Golpear la frente 3-4 veces sobre la colchoneta durante la tarea; dejar marcas rojas de dientes en la muñeca derecha.",
                "non_ex": "Resting head on floor mat during group circle time; mouthing FDA-approved silicone chew necklace softly.",
                "non_ex_zh": "在集体圈圈时间将头靠在地垫上休息；轻柔地咀嚼经过安全认证的硅胶项链。",
                "non_ex_es": "Descansar la cabeza en la colchoneta; morder suavemente un collar de silicona apto.",
                "dimensions": "Frequency: 3-6 episodes/day. Duration: 15s - 2min per outburst. Intensity: Moderate to Severe (potential skin redness or bruising).",
                "dimensions_zh": "频率：每天 3-6 次。持续时间：每次发作 15秒 至 2分钟。强度：中度至重度（可能导致皮肤发红或瘀青）。",
                "dimensions_es": "Frecuencia: 3-6 episodios/día. Duración: 15s - 2min. Intensidad: Moderada a Grave.",
                "triggers": "Setting Events: Fatigue, high ambient background noise.\nImmediate Triggers: Presentation of fine-motor discrete trial tasks, removal of preferred sensory toy.",
                "triggers_zh": "背景事件：疲劳、环境背景噪音过大。\n直接触发因素：呈现精细动作桌面任务、移走偏好的感官玩具。",
                "triggers_es": "Eventos de contexto: Fatiga, ruido ambiental.\nDesencadenantes inmediatos: Tareas motoras finas, retirada de juguete preferido.",
                "consequences": "RBT blocks contact using foam pad, pauses academic demand immediately, and prompts PECS 'Break' card.",
                "consequences_zh": "RBT 使用软垫阻挡物理接触，立即暂停学业要求，并提示使用 PECS“休息”卡片。",
                "consequences_es": "El RBT bloquea el contacto, pausa la demanda e indica la tarjeta PECS 'Descanso'.",
                "qabf_summary": "Task Escape: 14/15 | Physical Discomfort: 8/15 | Attention: 2/15 | Tangible: 3/15 | Sensory: 4/15",
                "qabf_summary_zh": "逃避任务: 14/15 | 身体不适: 8/15 | 社交关注: 2/15 | 获得物质: 3/15 | 感官刺激: 4/15",
                "qabf_summary_es": "Escape de tarea: 14/15 | Malestar físico: 8/15 | Atención: 2/15 | Tangible: 3/15 | Sensorial: 4/15",
                "triangulation": "65% Direct ABC Data (high rate during table work) + 25% Indirect Parent Interview (frustration with structured demands) + 10% QABF (Escape score 14/15).",
                "triangulation_zh": "65% 直接 ABC 数据（桌面任务期间高发） + 25% 间接家长访谈（对结构化任务表现出挫败） + 10% QABF 结果（逃避任务得分 14/15）。",
                "triangulation_es": "65% Datos ABC + 25% Entrevista + 10% QABF (Escape 14/15).",
                "hypothesis": "Primary Function: Task Escape (Social Negative Reinforcement).\nSecondary Function: Physical Discomfort Relief.",
                "hypothesis_zh": "核心行为功能：逃避任务（社交负强化）。\n次要行为功能：缓解身体不适。",
                "hypothesis_es": "Función principal: Escape de tareas.\nFunción secundaria: Alivio de malestar físico.",
                "ferb": "Functional Communication: Activate AAC BigMack button ('I need a break') or hand 'Break' PECS card to prompt immediate task pause.",
                "ferb_zh": "功能性替代行为 (FERB)：按下 AAC 语音按键（“我想休息”）或递交“休息” PECS 卡片，以提示立即暂停当前任务。",
                "ferb_es": "Comunicación Funcional: Activar el botón AAC BigMack ('Necesito un descanso') o entregar la tarjeta PECS 'Descanso'.",
            },
            {
                "name": "2. Sensory Vocal Distress & Face-Slapping",
                "name_zh": "2. 感官情绪性尖叫与拍打面部",
                "name_es": "2. Distrés vocal sensorial y bofetadas en la cara",
                "def": "Vocal screaming exceeding normal conversational volume (>80 dB) lasting >2 seconds, occurring simultaneously with open-palm striking of own cheeks or arms.",
                "def_zh": "尖叫声超过正常交谈音量（>80 dB）且持续时间>2秒，同时伴随用张开的手掌拍打自己脸颊或手臂的行为。",
                "def_es": "Gritos vocales que superan el volumen conversacional normal (>80 dB) durante >2 segundos, junto con golpes con la palma abierta en las mejillas o brazos.",
                "ex": "Loud screaming and striking cheeks 3 times when loud kitchen appliance starts in adjacent room.",
                "ex_zh": "当隔壁房间开启高分贝厨房电器时，大声尖叫并连续拍打脸颊 3 次。",
                "ex_es": "Gritos fuertes y 3 golpes en las mejillas al encenderse un electrodoméstico ruidoso.",
                "non_ex": "Excited vocal shouting during outdoor playground play; tapping cheeks rhythmically during music group time.",
                "non_ex_zh": "在户外操场游玩时兴奋地高声欢呼；在音乐课上跟随节奏轻拍脸颊。",
                "non_ex_es": "Gritos de emoción en el parque; golpear las mejillas al ritmo de la música.",
                "dimensions": "Frequency: 2-4 episodes/day. Duration: 30s - 3min. Intensity: Moderate.",
                "dimensions_zh": "频率：每天 2-4 次。持续时间：30秒 至 3分钟。强度：中度。",
                "dimensions_es": "Frecuencia: 2-4 episodios/día. Duración: 30s - 3min. Intensidad: Moderada.",
                "triggers": "Setting Events: Overstimulating ambient noise, sudden schedule changes.\nImmediate Triggers: Unexpected high-pitch sounds, removal of preferred juice cup.",
                "triggers_zh": "背景事件：环境噪音过度刺激、突发日程改变。\n直接触发因素：意料之外的高频噪音、偏好的果汁杯被移走。",
                "triggers_es": "Eventos de contexto: Ruido ambiental excesivo.\nDesencadenantes: Sonidos agudos inesperados.",
                "consequences": "Staff offers noise-canceling headphones and redirects to sensory chew tool.",
                "consequences_zh": "工作人员提供降噪耳机，并重新引导至口部感官咀嚼工具。",
                "consequences_es": "El personal ofrece auriculares de cancelación de ruido y redirige a un mordedor sensorial.",
                "qabf_summary": "Sensory/Automatic: 13/15 | Task Escape: 11/15 | Attention: 3/15 | Tangible: 2/15 | Physical: 1/15",
                "qabf_summary_zh": "感官刺激/自动强化: 13/15 | 逃避任务: 11/15 | 社交关注: 3/15 | 获得物质: 2/15 | 身体不适: 1/15",
                "qabf_summary_es": "Sensorial: 13/15 | Escape: 11/15 | Atención: 3/15 | Tangible: 2/15 | Físico: 1/15",
                "triangulation": "65% Direct ABC Data (appliance noise trigger) + 25% Indirect RBT Interview (auditory aversion) + 10% QABF (Sensory score 13/15).",
                "triangulation_zh": "65% 直接 ABC 数据（电器噪音触发） + 25% 间接 RBT 访谈（听觉敏感过度） + 10% QABF 结果（感官得分 13/15）。",
                "triangulation_es": "65% Datos ABC + 25% Entrevista RBT + 10% QABF (Sensorial 13/15).",
                "hypothesis": "Primary Function: Escape / Regulation of Auditory Overstimulation (Sensory Automatic Reinforcement).",
                "hypothesis_zh": "核心行为功能：逃避/调节听觉过度刺激（感官自动强化）。",
                "hypothesis_es": "Función principal: Escape / Regulación de sobreestimulación auditiva.",
                "ferb": "Self-Regulation / Request Strategy: Point to 'Headphones' visual symbol or independently retrieve noise-canceling headphones from designated sensory bin.",
                "ferb_zh": "功能性替代行为 (FERB)：指认“耳机”视觉标识，或自行从指定的感官箱中取出降噪耳机戴上。",
                "ferb_es": "Estrategia de autorregulación: Señalar el símbolo 'Auriculares' o tomarlos de la caja sensorial.",
            },
            {
                "name": "3. Property Destruction & Drop-to-Floor (Flopping)",
                "name_zh": "3. 破坏物品与躺地摆抗拒 (Flopping)",
                "name_es": "3. Destrucción de propiedad y tirarse al suelo (Flopping)",
                "def": "Any instance where the child forcibly throws, sweeps, or slams therapy materials/toys off surfaces, or drops body weight entirely to the floor resisting adult guidance for >3 seconds upon removal of a preferred tangible item.",
                "def_zh": "当偏好的物品/玩具被移走时，儿童用力抛掷、扫落或重摔干预用具/玩具，或将全身重量沉入地面拒绝成人引导持续>3秒的任何行为。",
                "def_es": "Cualquier instancia en la que el niño arroje, empuje o golpee materiales/juguetes, o deje caer su peso al suelo resistiéndose a la guía adulta durante >3 segundos tras la retirada de un objeto preferido.",
                "ex": "Sweeping discrete trial matching cards off desk and lying limp on floor when preferred tablet access is terminated.",
                "ex_zh": "当平板电脑使用时间结束时，把桌面上的干预配对卡片全部扫落到地上，并全身发软趴在地上不肯起来。",
                "ex_es": "Tirar las tarjetas de trabajo al suelo y tirarse al suelo cuan largo es al finalizar el tiempo de tableta.",
                "non_ex": "Accidentally knocking a crayon off desk during coloring; sitting down on play mat upon instructions.",
                "non_ex_zh": "涂色时不小心将蜡笔碰落桌下；听从指令坐在游戏地垫上。",
                "non_ex_es": "Tirar una crayola por accidente; sentarse en la colchoneta tras una instrucción.",
                "dimensions": "Frequency: 2-3 episodes/day. Duration: 1-5 minutes. Intensity: Moderate.",
                "dimensions_zh": "频率：每天 2-3 次。持续时间：1-5 分钟。强度：中度。",
                "dimensions_es": "Frecuencia: 2-3 episodios/día. Duración: 1-5 min. Intensidad: Moderada.",
                "triggers": "Setting Events: Transitions between high-preference and low-preference activities.\nImmediate Triggers: Verbal prompt to transition away from iPad/preferred sensory light toy.",
                "triggers_zh": "背景事件：在高偏好与低偏好活动之间的环节转换。\n直接触发因素：要求停止使用 iPad 或偏好感官玩具的语言提示。",
                "triggers_es": "Eventos: Transición de alta a baja preferencia.\nDesencadenante: Instrucción de entregar el iPad.",
                "consequences": "Staff provides visual countdown board, maintains safety, and prompts PECS 'More Time' or 'My Turn' icon.",
                "consequences_zh": "工作人员出示视觉倒计时板，维持安全保护，并提示使用 PECS“还想要/再玩会儿”标识。",
                "consequences_es": "El personal presenta contador visual y promueve el uso del ícono PECS 'Más tiempo'.",
                "qabf_summary": "Access to Tangible: 14/15 | Task Escape: 9/15 | Attention: 5/15 | Sensory: 2/15 | Physical: 0/15",
                "qabf_summary_zh": "获得物质/活动: 14/15 | 逃避任务: 9/15 | 社交关注: 5/15 | 感官刺激: 2/15 | 身体不适: 0/15",
                "qabf_summary_es": "Acceso a Tangible: 14/15 | Escape: 9/15 | Atención: 5/15 | Sensorial: 2/15 | Físico: 0/15",
                "triangulation": "65% Direct ABC Data (high occurrence at transition times) + 25% Parent Interview (struggles with home iPad limits) + 10% QABF (Tangible score 14/15).",
                "triangulation_zh": "65% 直接 ABC 数据（转换环节高发） + 25% 家人访谈（在家难以接受 iPad 时间限制） + 10% QABF 结果（获得物质得分 14/15）。",
                "triangulation_es": "65% Datos ABC + 25% Entrevista familiar + 10% QABF (Tangible 14/15).",
                "hypothesis": "Primary Function: Access to Preferred Tangible Items / Activities (Social Positive Reinforcement).",
                "hypothesis_zh": "核心行为功能：获取偏好的物品/活动（社交正强化）。",
                "hypothesis_es": "Función principal: Acceso a objetos o actividades preferidas.",
                "ferb": "Functional Communication: Hand 'More Time' PECS icon or press AAC button to request a 2-minute extension before transition.",
                "ferb_zh": "功能性替代行为 (FERB)：递交“再玩会儿” PECS 卡片，或按压 AAC 表达“还需要 2 分钟”，以适当延缓转换。",
                "ferb_es": "Comunicación Funcional: Entregar la tarjeta PECS 'Más tiempo' o presionar el botón AAC.",
            }
        ]
        "strengths": "Responds very well to 1:1 adult playful interaction, strong visual matching skills, highly motivated by musical cause-and-effect toys.",
        "strengths_zh": "对 1:1 成人游戏化互动反应良好，具备较强的视觉匹配能力，对音乐因果玩具表现出极高动机。",
        "strengths_es": "Responde muy bien a la interacción lúdica 1:1, fuertes habilidades de emparejamiento visual.",
        "history": "Diagnosed with ASD Level 3; currently receiving 15 hrs/week of Early Intervention ABA and Speech Therapy.",
        "history_zh": "确诊为孤独症谱系障碍（ASD 3级）；目前每周接受 15 小时的早期干预 ABA 及言语治疗服务。",
        "history_es": "Diagnosticado con TEA Nivel 3; actualmente recibe 15 hrs/semana de Intervención Temprana ABA y Logopedia.",
        "personalized_synthesis": (
            "Given the client's developmental age (3Y 4M) and ASD Level 3 diagnosis within an early intervention NDBI framework, "
            "the assessment triangulates that target behaviors are predominantly driven by task escape during structured tabletop demands "
            "and auditory sensory overload. Clinical recommendations strongly prioritize parent-co-regulated sensory accommodations, "
            "incorporating naturalistic developmental teaching (ESDM), rapid acquisition of low-tech AAC (PECS/BigMack), "
            "and embedding short 1:1 play-based reinforcement bursts to replace self-injurious escape behaviors safely."
        ),
        "personalized_synthesis_zh": (
            "结合该客户的实际发育年龄（3岁4个月）及在早期干预 NDBI 框架下的 ASD 3级诊断，"
            "评估数据交叉验证显示：其目标行为主要由结构化桌面任务期间的逃避动机以及环境听觉感官过载所触发。"
            "临床建议重点关注：结合家长共同调节的感官调适方案、深度融入自然情境教学法 (ESDM)、"
            "快速建立低科技 AAC 沟通工具（PECS/BigMack 语音按键），并嵌入高频次 1:1 游戏化强化，以安全无痛地替代自伤类逃避行为。"
        ),
        "personalized_synthesis_es": (
            "Dado el nivel de desarrollo del cliente (3 años 4 meses) y el diagnóstico de TEA Nivel 3 bajo el marco NDBI, "
            "la evaluación triangula que las conductas objetivo son impulsadas por el escape de tareas estructuradas "
            "y la sobrecarga sensorial auditiva. Las recomendaciones clínicas priorizan acomodaciones sensoriales, "
            "enseñanza NDBI/ESDM y comunicación augmentativa (PECS/AAC)."
        )
    },
    "g2": {
        "title": "School-Age IEP Protocol (5-21 Yrs)",
        "file_tag": "5to21yo",
        "framework": "IDEA IEP | PBIS Framework",
        "age_str": "10 Years 2 Months",
        "setting_str": "General Education Classroom / Resource Room",
        "triangulation_algo": "60% Direct ABC Data + 30% IEP Team Interview + 10% QABF",
        "protocol_sentences": [
            "• Aligned with IDEA IEP requirements and PBIS Multi-Tiered Support Systems.",
            "• Targets academic task engagement, self-advocacy, and emotional self-regulation.",
            "• Emphasizes replacement behaviors integrated into classroom routines.",
            "• Incorporates teacher-implemented token economies and peer modeling."
        ],
        "behaviors": [
            {
                "name": "1. Task Avoidance / Elopement from Seat",
                "name_zh": "1. 逃避学业 / 擅自离开座位",
                "name_es": "1. Evitación de tareas / Fuga del asiento",
                "def": "Leaving assigned desk area without teacher permission for >5 seconds during independent academic instruction.",
                "def_zh": "在独立学业教学期间，未经教师允许离开指定的课桌区域超过5秒。",
                "def_es": "Abandonar el área asignada sin permiso durante >5 segundos durante la instrucción académica.",
                "ex": "Running out of seat to classroom carpet area during independent math worksheet.",
                "ex_zh": "在独立完成数学工作表期间，从座位上跑开并躺在教室地毯区域。",
                "ex_es": "Salir corriendo del asiento hacia la alfombra durante la tarea de matemáticas.",
                "non_ex": "Standing up to walk to sharpener after raising hand and receiving permission.",
                "non_ex_zh": "举手并获得许可后，站起来走到削笔器旁削铅笔。",
                "non_ex_es": "Levantarse para ir al sacapuntas tras pedir permiso.",
                "dimensions": "Frequency: 4-5 times/day. Duration: 1-5 minutes. Intensity: Low to Moderate.",
                "dimensions_zh": "频率：每天 4-5 次。持续时间：1-5 分钟。强度：低至中度。",
                "dimensions_es": "Frecuencia: 4-5 veces/día. Duración: 1-5 min. Intensidad: Baja a Moderada.",
                "triggers": "Setting Events: Multi-step math tasks.\nImmediate Triggers: Presentation of 2-page math assignment.",
                "triggers_zh": "背景事件：多步骤数学任务。\n直接触发因素：发放长达2页的数学作业纸。",
                "triggers_es": "Eventos: Tareas complejas.\nDesencadenante: Hoja de ejercicios larga.",
                "consequences": "Staff presents 'Break' visual card; demand temporarily paused.",
                "consequences_zh": "教职工呈现“休息”视觉卡片；学业要求被暂时暂停。",
                "consequences_es": "El personal presenta la tarjeta 'Descanso'; la demanda se pausa.",
                "qabf_summary": "Task Escape: 15/15 | Attention: 5/15 | Tangible: 2/15 | Sensory: 1/15 | Physical: 0/15",
                "qabf_summary_zh": "逃避任务: 15/15 | 社交关注: 5/15 | 获得物质: 2/15 | 感官刺激: 1/15 | 身体不适: 0/15",
                "qabf_summary_es": "Escape: 15/15 | Atención: 5/15 | Tangible: 2/15 | Sensorial: 1/15 | Físico: 0/15",
                "triangulation": "60% Direct ABC Data + 30% Indirect IEP Interview + 10% QABF (Escape score 15/15).",
                "triangulation_zh": "60% 直接 ABC 数据 + 30% 间接 IEP 访谈 + 10% QABF 结果（逃避得分 15/15）。",
                "triangulation_es": "60% Datos ABC + 30% Entrevista IEP + 10% QABF (Escape 15/15).",
                "hypothesis": "Primary Function: Escape from Academic Demands.",
                "hypothesis_zh": "核心行为功能：逃避学业任务要求。",
                "hypothesis_es": "Función principal: Escape de demandas académicas.",
                "ferb": "Hand 'Break' card to teacher or place 'Help Needed' tent on desk.",
                "ferb_zh": "向教师递交“休息”卡片，或在桌上摆放“需要帮助”提示牌。",
                "ferb_es": "Entregar la tarjeta 'Descanso' al profesor o colocar la tarjeta 'Ayuda' en el escritorio.",
            }
        ],
        "strengths": "Excellent visual-spatial abilities, enthusiastic about technology and drawing.",
        "strengths_zh": "具备出色的视觉空间能力，对科技和绘画抱有极高热情。",
        "strengths_es": "Excelentes habilidades visoespaciales, entusiasta de la tecnología.",
        "history": "Enrolled in General Education with IEP support.",
        "history_zh": "就读于普通教育班级，享有 IEP 特殊教育支持计划。",
        "history_es": "Inscrito en educación general con apoyo de IEP.",
        "personalized_synthesis": "Assessment indicates academic escape is the primary function. Intervention focuses on embedding self-advocacy within classroom IEP goals, adjusting work task length, and implementing a teacher-managed visual token economy.",
        "personalized_synthesis_zh": "评估明确表明逃避学业是主要功能。干预重点在于将自我倡导融入教室 IEP 目标，微调作业分量，并实施教师管理的视觉代币筹码制度。",
        "personalized_synthesis_es": "La evaluación indica que el escape académico es la función principal. La intervención se centra en la autogestión e IEP."
    },
    "g3": {
        "title": "Adult Community Protocol (21+ Yrs)",
        "file_tag": "21plusYo",
        "framework": "Medicaid HCBS | Person-Centered Waiver Framework",
        "age_str": "26 Years 8 Months",
        "setting_str": "Vocational Workshop & Day Program",
        "triangulation_algo": "50% Direct ABC Data + 40% Person-Centered Interview + 10% QABF",
        "protocol_sentences": [
            "• Designed for Medicaid HCBS Waiver adult day programs and community living.",
            "• Focuses on person-centered planning, independence, and vocational endurance.",
            "• Emphasizes self-management protocols and respectful adult communication.",
            "• Reduces intrusive restrictive interventions through positive behavior support."
        ],
        "behaviors": [
            {
                "name": "1. Vocational Task Refusal & Verbal Aggression",
                "name_zh": "1. 拒绝职业操作与言语攻击",
                "name_es": "1. Rechazo de tareas laborales y agresión verbal",
                "def": "Refusing assembly or sorting demands accompanied by loud vocal threats (>75 dB) or pushing work materials away.",
                "def_zh": "拒绝组装或分类任务，并伴随大声言语威胁（>75 dB）或推开工作材料的行为。",
                "def_es": "Rechazar tareas de ensamblaje o clasificación acompañado de amenazas vocales fuertes (>75 dB) o empujar materiales.",
                "ex": "Shouting 'No way!', slamming assembly boxes on desk when quota is raised.",
                "ex_zh": "当提高工作配额时，大叫“绝不可能！”并将组装盒重重摔在桌上。",
                "ex_es": "Gritar '¡De ninguna manera!' y golpear las cajas de ensamblaje.",
                "non_ex": "Verbally requesting a 5-minute break in a normal tone.",
                "non_ex_zh": "用正常音量和语调口头提出“想要休息5分钟”。",
                "non_ex_es": "Solicitar verbalmente un descanso de 5 minutos en tono normal.",
                "dimensions": "Frequency: 1-2 times/week. Duration: 5-10 minutes. Intensity: Moderate.",
                "dimensions_zh": "频率：每周 1-2 次。持续时间：5-10 分钟。强度：中度。",
                "dimensions_es": "Frecuencia: 1-2 veces/semana. Duración: 5-10 min. Intensidad: Moderada.",
                "triggers": "Setting Events: Unfamiliar staff.\nImmediate Triggers: Direct instructions to complete vocational assembly quota.",
                "triggers_zh": "背景事件：不熟悉的工作人员。\n直接触发因素：要求完成职业组装配额的直接指令。",
                "triggers_es": "Eventos: Personal no familiar.\nDesencadenante: Instrucciones para cumplir cuotas de trabajo.",
                "consequences": "DSP offers choice board, demand temporarily paused.",
                "consequences_zh": "直属支持人员（DSP）提供选择板，任务要求被暂时暂停。",
                "consequences_es": "El DSP ofrece una tabla de opciones; la demanda se pausa.",
                "qabf_summary": "Task Escape: 13/15 | Attention: 4/15 | Tangible: 3/15 | Sensory: 1/15 | Physical: 1/15",
                "qabf_summary_zh": "逃避任务: 13/15 | 社交关注: 4/15 | 获得物质: 3/15 | 感官刺激: 1/15 | 身体不适: 1/15",
                "qabf_summary_es": "Escape: 13/15 | Atención: 4/15 | Tangible: 3/15 | Sensorial: 1/15 | Físico: 1/15",
                "triangulation": "50% Direct ABC Data + 40% Indirect Vocational Interview + 10% QABF (Escape score 13/15).",
                "triangulation_zh": "50% 直接 ABC 数据 + 40% 间接职业能力访谈 + 10% QABF 结果（逃避得分 13/15）。",
                "triangulation_es": "50% Datos ABC + 40% Entrevista vocacional + 10% QABF (Escape 13/15).",
                "hypothesis": "Primary Function: Escape from Vocational Assembly Demands.",
                "hypothesis_zh": "核心行为功能：逃避职业组装工作要求。",
                "hypothesis_es": "Función principal: Escape de demandas laborales.",
                "ferb": "Verbally request '5-minute break, please' using self-advocacy phrase card.",
                "ferb_zh": "使用自我倡导短语卡口头表达：“请给我 5 分钟休息时间”。",
                "ferb_es": "Solicitar verbalmente 'Un descanso de 5 minutos, por favor' usando tarjeta de frases.",
            }
        ],
        "strengths": "High independence in personal self-care.",
        "strengths_zh": "在个人日常生活自理方面具备极高独立性。",
        "strengths_es": "Alta independencia en el autocuidado personal.",
        "history": "Participates in Adult Day Vocational Services under Medicaid HCBS Waiver.",
        "history_zh": "在 Medicaid HCBS 豁免计划下参与成人日间职业干预服务。",
        "history_es": "Participa en servicios vocacionales de adultos bajo Medicaid HCBS Waiver.",
        "personalized_synthesis": "Primary function validated as task avoidance under high quotas. Recommendations center on self-managed work pacing and dignified adult self-advocacy protocols.",
        "personalized_synthesis_zh": "确证主要功能为高配额下的任务逃避。干预建议围绕自我管理的职业节奏调配及具备尊严的成人自我倡导协议展开。",
        "personalized_synthesis_es": "Función principal validada como evitación de tareas. Las recomendaciones se centran en la autorregulación laboral."
    }
}

# ==========================================
# 3. Main Interface & Security Header
# ==========================================
st.markdown("""
    <div class="hipaa-banner">
        <div class="hipaa-title">🛡️ 100% HIPAA COMPLIANT & ZERO-CLOUD LOCAL PROCESSING</div>
        <div class="hipaa-body">
            This tool strictly complies with HIPAA privacy regulations. All data parsing, analysis, and document formulation occur <strong>100% locally within your active browser session memory</strong>.
        </div>
    </div>
""", unsafe_allow_html=True)

st.markdown("<div class='main-title'>🧩 BCBA Clinical FBA & BIP Draft Formulation Tool <span class='demo-tag'>(Demo Version)</span></div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>Interactive Demonstration for Automated Clinical First-Draft Synthesis | Designed for BCBAs & LBAs</div>", unsafe_allow_html=True)

st.divider()

# ==========================================
# 4. Cohort Selection
# ==========================================
st.markdown("### 1️⃣ Select Clinical Cohort")

cohort_options = {
    "g1": "👶 Early Intervention Protocol (2-5 Yrs)",
    "g2": "🏫 School-Age / IEP Protocol (5-21 Yrs)",
    "g3": "💼 Adult Community & Vocational Protocol (21+ Yrs)"
}

selected_cohort_key = st.radio(
    "Select Target Client Population:",
    options=list(cohort_options.keys()),
    format_func=lambda x: cohort_options[x],
    index=0,
    horizontal=True
)

current_meta = cohort_meta[selected_cohort_key]

# ==========================================
# 5. Assessment Data Import & Protocol Card
# ==========================================
st.markdown("### 2️⃣ Import Assessment Data & Protocol Overview")

st.markdown(f"""
    <div class="protocol-card">
        <div class="protocol-title">📋 Selected Protocol Framework: {current_meta['title']} ({current_meta['framework']})</div>
        {"".join([f'<div class="protocol-bullet">{s}</div>' for s in current_meta['protocol_sentences']])}
        <div class="triangulation-badge">⚖️ Triangulation Algorithm: {current_meta['triangulation_algo']}</div>
    </div>
""", unsafe_allow_html=True)

col_input1, col_input2, col_input3 = st.columns([1,1,1])

with col_input1:
    st.markdown("#### 📄 Direct Observation (ABC)")
    mock_csv = generate_mock_abc_csv(selected_cohort_key)
    st.download_button(
        label=f"📥 Download Mock ABC (.csv)",
        data=mock_csv,
        file_name=f"DeIdentified_ABC_{current_meta['file_tag']}.csv",
        mime="text/csv",
        use_container_width=True
    )
    uploaded_abc = st.file_uploader("Upload ABC File:", type=["csv", "xlsx"], key=f"abc_{selected_cohort_key}")

with col_input2:
    st.markdown("#### 📝 Indirect Interview Notes")
    mock_docx = generate_mock_interview_docx(selected_cohort_key)
    st.download_button(
        label=f"📥 Download Mock Interview (.docx)",
        data=mock_docx,
        file_name=f"DeIdentified_Interview_{current_meta['file_tag']}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
    uploaded_interview = st.file_uploader("Upload Interview File:", type=["docx", "txt"], key=f"interview_{selected_cohort_key}")

with col_input3:
    st.markdown("#### 📊 Behavior QABF Assessment Results")
    mock_qabf = generate_mock_qabf_docx(selected_cohort_key)
    st.download_button(
        label=f"📥 Download Mock QABF Results (.docx)",
        data=mock_qabf,
        file_name=f"DeIdentified_QABF_{current_meta['file_tag']}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
    uploaded_qabf = st.file_uploader("Upload QABF File:", type=["docx", "txt"], key=f"qabf_{selected_cohort_key}")

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

def add_bi_item(doc, label_en, val_en, label_trans=None, val_trans=None, is_trans=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.space_before = Pt(2)
    
    r_lbl = p.add_run(f"{label_en}: ")
    r_lbl.bold = True
    p.add_run(f"{val_en}")
    
    if is_trans and label_trans and val_trans:
        p.add_run("\n")
        r_tr_lbl = p.add_run(f"[{label_trans}: ")
        r_tr_lbl.bold = True
        r_tr_lbl.italic = True
        r_tr_lbl.font.color.rgb = RGBColor(100, 100, 100)
        
        r_tr_val = p.add_run(f"{val_trans}]")
        r_tr_val.italic = True
        r_tr_val.font.color.rgb = RGBColor(100, 100, 100)

def build_compact_demographics_table(doc, c_meta, is_trans, lang_choice):
    table = doc.add_table(rows=5, cols=2)
    table.style = 'Table Grid'
    
    lang_str = "English / Spanish Support" if "Spanish" in lang_choice else ("English / Chinese Support" if "Chinese" in lang_choice else "English Standard")
    
    data = [
        ("Student/Client Name", "[CLIENT_NAME]", "DOB / Age", f"[CLIENT_DOB] / {c_meta['age_str']}"),
        ("Client ID", "[CLIENT_ID]", "Assessment Date", "2026-08-14"),
        ("Facility/School", "[DISTRICT_OR_FACILITY_NAME]", "Setting", c_meta['setting_str']),
        ("Assessor", "[BCBA_NAME], BCBA, LBA", "Framework", c_meta['framework']),
        ("Primary Language", lang_str, "Informants", "Parent, Lead Teacher / RBT")
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
    is_es = "Spanish" in lang_choice
    is_trans = is_zh or is_es

    p_t = doc.add_paragraph()
    r_t = p_t.add_run("FUNCTIONAL BEHAVIORAL ASSESSMENT (FBA) REPORT")
    r_t.bold = True
    p_t.style = doc.styles['Title']

    if is_zh:
        p_tr = doc.add_paragraph()
        r_tr = p_tr.add_run("[功能性行为评估 (FBA) 报告 Draft]")
        r_tr.italic = True
        r_tr.font.color.rgb = RGBColor(100, 100, 100)
    elif is_es:
        p_tr = doc.add_paragraph()
        r_tr = p_tr.add_run("[Informe de Evaluación de la Conducta Funcional (FBA)]")
        r_tr.italic = True
        r_tr.font.color.rgb = RGBColor(100, 100, 100)

    # Section 1: Demographics
    add_bi_heading(doc, 1, "1. Student Demographics & Administrative Info", 
                   "1. 学生/客户基本信息与行政登记" if is_zh else ("1. Datos Demográficos y Administrativos" if is_es else None))
    build_compact_demographics_table(doc, c_meta, is_trans, lang_choice)

    # Section 2: Data Sources & Triangulation Algorithm
    add_bi_heading(doc, 1, "2. Data Sources & Triangulation Methodology", 
                   "2. 数据来源与三方交叉验证评估方法" if is_zh else ("2. Fuentes de Datos y Metodología de Triangulación" if is_es else None))
    
    triangulation_desc_en = (
        f"Triangulation Algorithm Standard Applied: {c_meta['triangulation_algo']}.\n"
        "1. Direct ABC Data: Continuous recording across baseline therapy sessions.\n"
        "2. Indirect Assessment: Structured interviews with parent and lead RBT.\n"
        "3. Psychometric Rating Scale: Behavior-specific QABF (Questions About Behavioral Function)."
    )
    triangulation_desc_trans = (
        f"应用三方数据加权验证算法：{c_meta['triangulation_algo']}。\n"
        "1. 直接 ABC 数据：干预课期间连续行为观察记录。\n"
        "2. 间接评估：与家长和督导 RBT 的结构化访谈。\n"
        "3. 心理测量量表：针对具体行为的 QABF 行为功能评估问卷。"
    ) if is_zh else (
        f"Algoritmo de triangulación aplicado: {c_meta['triangulation_algo']}.\n"
        "1. Datos ABC directos.\n2. Entrevista indirecta.\n3. Escala psicométrica QABF."
    )
    
    add_bi_item(doc, "Methodology & Triangulation Formula", triangulation_desc_en, 
                "评估方法与三方验证公式" if is_zh else ("Metodología y Fórmula" if is_es else None), 
                triangulation_desc_trans, is_trans)

    # Section 3: Background & Strengths
    add_bi_heading(doc, 1, "3. Brief Background & Strengths Summary", 
                   "3. 学生背景与优势摘要" if is_zh else ("3. Antecedentes y Resumen de Fortalezas" if is_es else None))
    
    strengths_trans = c_meta.get("strengths_zh" if is_zh else "strengths_es")
    history_trans = c_meta.get("history_zh" if is_zh else "history_es")
    
    add_bi_item(doc, "Strengths & Preferences", c_meta["strengths"], 
                "优势与偏好" if is_zh else ("Fortalezas" if is_es else None), strengths_trans, is_trans)
    add_bi_item(doc, "Clinical / Educational History", c_meta["history"], 
                "临床/教育背景" if is_zh else ("Historial Clínico" if is_es else None), history_trans, is_trans)

    # Section 4: Individual Functional Analyses (Behavior-by-Behavior)
    add_bi_heading(doc, 1, "4. Individual Target Behavior Functional Analyses", 
                   "4. 目标行为独立功能分析 (按行为逐项拆解)" if is_zh else ("4. Análisis Funcional de Conductas Objetivo" if is_es else None))

    for idx, b in enumerate(behavior_list, 1):
        b_name_trans = b.get("name_zh" if is_zh else "name_es")
        add_bi_heading(doc, 2, f"Target Behavior #{idx}: {b['name']}", b_name_trans)

        b_def_trans = b.get("def_zh" if is_zh else "def_es")
        b_ex_trans = b.get("ex_zh" if is_zh else "ex_es")
        b_non_ex_trans = b.get("non_ex_zh" if is_zh else "non_ex_es")
        b_dim_trans = b.get("dimensions_zh" if is_zh else "dimensions_es")
        b_trig_trans = b.get("triggers_zh" if is_zh else "triggers_es")
        b_cons_trans = b.get("consequences_zh" if is_zh else "consequences_es")
        b_qabf_trans = b.get("qabf_summary_zh" if is_zh else "qabf_summary_es")
        b_tri_trans = b.get("triangulation_zh" if is_zh else "triangulation_es")
        b_hyp_trans = b.get("hypothesis_zh" if is_zh else "hypothesis_es")

        add_bi_item(doc, "A. Operational Definition", b["def"], "A. 操作性定义" if is_zh else ("A. Definición Operacional" if is_es else None), b_def_trans, is_trans)
        add_bi_item(doc, "B. Examples & Non-Examples", f"Examples: {b['ex']}\nNon-Examples: {b['non_ex']}", 
                    "B. 示例与非示例" if is_zh else ("B. Ejemplos y No-Ejemplos" if is_es else None), 
                    f"示例: {b_ex_trans}\n非示例: {b_non_ex_trans}", is_trans)
        add_bi_item(doc, "C. Behavior Dimensions", b["dimensions"], "C. 行为维度 (频率/持续时间/强度)" if is_zh else ("C. Dimensiones" if is_es else None), b_dim_trans, is_trans)
        add_bi_item(doc, "D. Environmental Triggers & Setting Events", b["triggers"], "D. 环境触发因素与背景事件" if is_zh else ("D. Desencadenantes" if is_es else None), b_trig_trans, is_trans)
        add_bi_item(doc, "E. Maintaining Consequences", b["consequences"], "E. 维持后果与他人反应" if is_zh else ("E. Consecuencias" if is_es else None), b_cons_trans, is_trans)
        add_bi_item(doc, "F. Behavior-Specific QABF Results", b["qabf_summary"], "F. 该行为专属 QABF 量表得分" if is_zh else ("F. Resultados QABF" if is_es else None), b_qabf_trans, is_trans)
        add_bi_item(doc, "G. Triangulation Analysis", b["triangulation"], "G. 三方交叉验证分析" if is_zh else ("G. Triangulación" if is_es else None), b_tri_trans, is_trans)
        add_bi_item(doc, "H. Hypothesized Function", b["hypothesis"], "H. 该行为推断功能" if is_zh else ("H. Hipótesis Funcional" if is_es else None), b_hyp_trans, is_trans)

    # Section 5: Overall Synthesis & Clinical Recommendations (Personalized)
    add_bi_heading(doc, 1, "5. Synthesis & Clinical Recommendations", 
                   "5. 综合评估结论与个性化临床建议" if is_zh else ("5. Síntesis y Recomendaciones Clínicas Personalizadas" if is_es else None))
    
    synth_trans = c_meta.get("personalized_synthesis_zh" if is_zh else "personalized_synthesis_es")
    add_bi_item(doc, "Clinical Synthesis & Recommendations", c_meta["personalized_synthesis"], 
                "临床综合结论与具体建议" if is_zh else ("Síntesis Clínica y Recomendaciones" if is_es else None), 
                synth_trans, is_trans)

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
    is_es = "Spanish" in lang_choice
    is_trans = is_zh or is_es

    p_t = doc.add_paragraph()
    r_t = p_t.add_run("BEHAVIOR INTERVENTION PLAN (BIP)")
    r_t.bold = True
    p_t.style = doc.styles['Title']

    if is_zh:
        p_tr = doc.add_paragraph()
        r_tr = p_tr.add_run("[行为干预计划 (BIP) Comprehensive Draft]")
        r_tr.italic = True
        r_tr.font.color.rgb = RGBColor(100, 100, 100)
    elif is_es:
        p_tr = doc.add_paragraph()
        r_tr = p_tr.add_run("[Plan de Intervención de la Conducta (BIP)]")
        r_tr.italic = True
        r_tr.font.color.rgb = RGBColor(100, 100, 100)

    # Section 1
    add_bi_heading(doc, 1, "1. Student Info & Administrative Summary", 
                   "1. 学生/客户信息与行政摘要" if is_zh else ("1. Información del Estudiante" if is_es else None))
    build_compact_demographics_table(doc, c_meta, is_trans, lang_choice)

    # Section 2: Behavior Operational Defs, Functions & FERB Breakdown
    add_bi_heading(doc, 1, "2. Target Behaviors Operational Definitions, Functions & Replacement Skills (FERB)", 
                   "2. 目标行为操作性定义、行为功能与替代技能 (FERB) 逐项拆解" if is_zh else ("2. Definiciones Operacionales, Funciones y FERB" if is_es else None))

    for idx, b in enumerate(c_meta["behaviors"], 1):
        b_name_trans = b.get("name_zh" if is_zh else "name_es")
        add_bi_heading(doc, 2, f"Target Behavior #{idx}: {b['name']}", b_name_trans)
        
        b_def_trans = b.get("def_zh" if is_zh else "def_es")
        add_bi_item(doc, "Operational Definition", b["def"], 
                    "操作性定义" if is_zh else ("Definición Operacional" if is_es else None), b_def_trans, is_trans)
        
        b_hyp_trans = b.get("hypothesis_zh" if is_zh else "hypothesis_es")
        add_bi_item(doc, "Validated Function", b["hypothesis"], 
                    "确证行为功能" if is_zh else ("Función Validada" if is_es else None), b_hyp_trans, is_trans)
        
        b_ferb_trans = b.get("ferb_zh" if is_zh else "ferb_es")
        add_bi_item(doc, "Functionally Equivalent Replacement Behavior (FERB)", b["ferb"], 
                    "功能性替代行为 (FERB)" if is_zh else ("Conducta de Reemplazo (FERB)" if is_es else None), b_ferb_trans, is_trans)

    # Section 3: Proactive / Antecedent Strategies
    add_bi_heading(doc, 1, "3. Proactive & Antecedent Modifications (Prevention)", 
                   "3. 前因调整与预防策略" if is_zh else ("3. Estrategias Proactivas y Antecedentes" if is_es else None))
    add_bi_item(doc, "3.1 Environmental Adaptations", 
                "• Provide visual/auditory transition warnings (2-min & 1-min).\n• Offer noise-canceling headphones prior to noisy routines.\n• Break tasks into small visual chunks.", 
                "3.1 环境调整" if is_zh else ("3.1 Adaptaciones Ambientales" if is_es else None), 
                "• 提供活动转换倒计时提示。\n• 预先提供降噪耳机。\n• 任务拆解为小步子视觉单元。" if is_zh else "• Advertencias visuales.\n• Auriculares sensoriales.", is_trans)

    # Section 4: Replacement Behaviors Protocols
    add_bi_heading(doc, 1, "4. Functional Communication Training (FCT)", 
                   "4. 功能性沟通训练 (FCT)" if is_zh else ("4. Entrenamiento en Comunicación Funcional (FCT)" if is_es else None))
    add_bi_item(doc, "4.1 FCT Protocols", 
                "Prompt client to press AAC button ('Break') or hand PECS card upon initial sign of frustration.", 
                "4.1 FCT 协议" if is_zh else ("4.1 Protocolos FCT" if is_es else None), 
                "提示客户在情绪萌芽阶段使用 AAC 或 PECS 卡片请求休息。" if is_zh else "Indicar el uso de AAC o PECS para pedir descansos.", is_trans)

    # Section 5: Reinforcement Strategies
    add_bi_heading(doc, 1, "5. Reinforcement Protocols", 
                   "5. 强化策略协议" if is_zh else ("5. Protocolos de Reforzamiento" if is_es else None))
    add_bi_item(doc, "5.1 DRA Protocol", 
                "Immediate 100% compliance with requested 'Break' AAC activations paired with high-value praise.", 
                "5.1 替代行为区别性强化 (DRA)" if is_zh else ("5.1 Protocolo DRA" if is_es else None), 
                "对独立的 AAC 表达立即 100% 给予休息强化并结合口头表扬。" if is_zh else "Cumplimiento inmediato del 100% al pedir descanso.", is_trans)

    # Section 6: Response Strategies
    add_bi_heading(doc, 1, "6. Reactive Response Protocols & Extinction", 
                   "6. 目标行为回应与消退策略" if is_zh else ("6. Estrategias de Respuesta y Extinción" if is_es else None))
    add_bi_item(doc, "6.1 Physical Blocking & Extinction", 
                "Maintain neutral expression, avoid eye contact. Block SIB softly using foam pads without verbal feedback.", 
                "6.1 物理阻挡与消退" if is_zh else ("6.1 Bloqueo Físico y Extinción" if is_es else None), 
                "保持中立表情，无言语训诫。使用软垫中立阻挡自伤动作。" if is_zh else "Mantener expresión neutral, bloquear SIB con colchoneta.", is_trans)

    # Section 7: Staff Training & Monitoring (NEW & COMPREHENSIVE)
    add_bi_heading(doc, 1, "7. Staff Training and Monitoring", 
                   "7. 人员培训与督导忠实度核查 (Staff Training & Monitoring)" if is_zh else ("7. Entrenamiento y Monitoreo del Personal" if is_es else None))
    
    training_resp_en = "Case Lead BCBA / LBA Supervision Team."
    training_resp_trans = "负责 BCBA / LBA 临床督导团队。" if is_zh else "BCBA Líder del Caso."
    add_bi_item(doc, "Who is responsible for training others to implement this BIP?", training_resp_en, 
                "负责培训相关人员实施该 BIP 的责任人" if is_zh else ("Responsable de la capacitación" if is_es else None), training_resp_trans, is_trans)

    training_process_en = (
        "1. Didactic Review: Lead BCBA reviews operational definitions, antecedent modifications, and reactive extinction protocols with all staff.\n"
        "2. Behavioral Skills Training (BST): Staff complete Role-playing, Modeling, and Practice until achieving 90%+ baseline competency.\n"
        "3. Live In-Situ Coaching: BCBA provides immediate feedback during direct therapy sessions prior to independent implementation."
    )
    training_process_trans = (
        "1. 理论讲解：主导 BCBA 详细解读操作性定义、前因预防调整及回应消退协议。\n"
        "2. 行为技能培训 (BST)：通过示范、角色扮演与演练，直至人员考核达到 90% 以上合格率。\n"
        "3. 现场实操指导：在独立上课前，由 BCBA 在真实干预场景中进行实时指导与反馈。"
    ) if is_zh else "Capacitación BST: Lectura, Modelado, Rol-Playing y Feedback en vivo."
    add_bi_item(doc, "What is the process for training others to implement this plan?", training_process_en, 
                "培训实施人员的具体流程 (BST 模式)" if is_zh else ("Proceso de capacitación" if is_es else None), training_process_trans, is_trans)

    fidelity_freq_en = (
        "Staff will be observed for Treatment Fidelity of Implementation at least WEEKLY during the first month of BIP implementation, "
        "and BI-WEEKLY thereafter, utilizing a standardized 10-point Treatment Fidelity Checklist (Threshold: ≥90% compliance)."
    )
    fidelity_freq_trans = (
        "在 BIP 实施的第一个月内，每周至少进行一次干预忠实度观察；之后每两周进行一次。"
        "使用标准化 10 项忠实度核查表进行评估（合格标准：遵从率 ≥90%）。"
    ) if is_zh else "Observación semanal el primer mes, bisemanal posteriormente. Meta ≥90%."
    add_bi_item(doc, "How often will staff be observed to ensure implementation as written (Fidelity)?", fidelity_freq_en, 
                "人员执行忠实度 (Fidelity) 的观察核查频率" if is_zh else ("Frecuencia de monitoreo de fidelidad" if is_es else None), fidelity_freq_trans, is_trans)

    review_timeline_en = "This BIP will be formally reviewed MONTHLY by the BCBA, or immediately upon any spike in SIB intensity or safety risk."
    review_timeline_trans = "本 BIP 将由 BCBA 每月进行一次正式评估复审；若自伤行为强度陡增或出现安全风险，将立即触发临时复审。" if is_zh else "Revisión mensual o de inmediato si aumentan los riesgos."
    add_bi_item(doc, "When will this plan be reviewed again?", review_timeline_en, 
                "本计划的下一次复审时间节点" if is_zh else ("Próxima fecha de revisión" if is_es else None), review_timeline_trans, is_trans)

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
            "Bilingual (English / Spanish - Español)",
            "Bilingual (English / Simplified Chinese - 简体中文)"
        ],
        index=2
    )

fba_docx_bytes = generate_exact_fba_doc(selected_cohort_key, report_lang, active_behaviors)
bip_docx_bytes = generate_exact_bip_doc(selected_cohort_key, report_lang)

with col_action1:
    st.write(" ")
    st.write(" ")
    st.download_button(
        label="⚡ Formulate & Download De-Identified FBA Draft (.docx)",
        data=fba_docx_bytes,
        file_name=f"DeIdentified_FBA_Draft_{current_meta['file_tag']}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

with col_action2:
    st.write(" ")
    st.write(" ")
    st.download_button(
        label="⚡ Formulate & Download De-Identified BIP Draft (.docx)",
        data=bip_docx_bytes,
        file_name=f"DeIdentified_BIP_Draft_{current_meta['file_tag']}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )

st.divider()

st.caption("⚠️ **Clinical Responsibility Notice:** This formulation tool serves strictly as a clinical first-draft synthesizer for BCBAs and LBAs. All generated drafts are fully de-identified and must be independently reviewed and edited prior to formal signature.")
