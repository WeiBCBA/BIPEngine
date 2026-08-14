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
                "• Emphasizes proactive environmental adaptation and rapid"
                " reinforcement for replacement skills."
            ),
        ],
        "behaviors": [
            {
                "name": (
                    "1. Self-Injurious Behavior (SIB) - Head Banging & Wrist"
                    " Biting"
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
                "ex": (
                    "Banging forehead 3-4 times on foam mat during tabletop"
                    " task; leaving red teeth marks on right wrist during transition."
                ),
                "ex_zh": (
                    "在桌面任务期间，额头在泡棉垫上连续碰撞3-4次；在转换环节中在右手腕上留下红印牙痕。"
                ),
                "non_ex": (
                    "Resting head on floor mat during group circle time;"
                    " mouthing FDA-approved silicone chew necklace softly."
                ),
                "non_ex_zh": (
                    "在集体圈圈时间将头靠在地垫上休息；轻柔地咀嚼经过安全认证的硅胶项链。"
                ),
                "dimensions": (
                    "Frequency: 3-6 episodes/day. Duration: 15s - 2min per"
                    " outburst. Intensity: Moderate to Severe (potential skin"
                    " redness or bruising)."
                ),
                "dimensions_zh": (
                    "频率：每天 3-6 次。持续时间：每次发作 15秒 至 2分钟。强度：中度至重度（可能导致皮肤发红或瘀青）。"
                ),
                "triggers": (
                    "Setting Events: Fatigue, high ambient background noise.\nImmediate"
                    " Triggers: Presentation of fine-motor discrete trial"
                    " tasks, removal of preferred sensory toy."
                ),
                "triggers_zh": (
                    "背景事件：疲劳、环境背景噪音过大。\n直接触发因素：呈现精细动作桌面任务、移走偏好的感官玩具。"
                ),
                "consequences": (
                    "RBT blocks contact using foam pad, pauses academic demand"
                    " immediately, and prompts PECS 'Break' card."
                ),
                "consequences_zh": (
                    "RBT 使用软垫阻挡物理接触，立即暂停学业要求，并提示使用 PECS“休息”卡片。"
                ),
                "qabf_summary": (
                    "Task Escape: 14/15 | Physical Discomfort: 8/15 |"
                    " Attention: 2/15 | Tangible: 3/15 | Sensory: 4/15"
                ),
                "qabf_summary_zh": (
                    "逃避任务: 14/15 | 身体不适: 8/15 | 社交关注: 2/15 | 获得物质: 3/15 |"
                    " 感官刺激: 4/15"
                ),
                "triangulation": (
                    "Direct ABC Data (high rate during table work) + Indirect"
                    " Parent Interview (frustration with structured demands) +"
                    " QABF (Escape score 14/15)."
                ),
                "triangulation_zh": (
                    "直接 ABC 数据（桌面任务期间高发） + 间接家长访谈（对结构化任务表现出挫败） +"
                    " QABF 结果（逃避任务得分 14/15）。"
                ),
                "hypothesis": (
                    "Primary Function: Task Escape (Social Negative"
                    " Reinforcement).\nSecondary Function: Physical Discomfort"
                    " Relief."
                ),
                "hypothesis_zh": "核心行为功能：逃避任务（社交负强化）。\n次要行为功能：缓解身体不适。",
                "ferb": (
                    "Functional Communication: Activate AAC BigMack button"
                    " ('I need a break') or hand 'Break' PECS card to prompt"
                    " immediate task pause."
                ),
                "ferb_zh": (
                    "功能性替代行为 (FERB)：按下 AAC 语音按键（“我想休息”）或递交“休息”"
                    " PECS 卡片，以提示立即暂停当前任务。"
                ),
            },
            {
                "name": "2. Sensory Vocal Distress & Face-Slapping",
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
                "ex": (
                    "Loud screaming and striking cheeks 3 times when loud"
                    " kitchen appliance starts in adjacent room."
                ),
                "ex_zh": "当隔壁房间开启高分贝厨房电器时，大声尖叫并连续拍打脸颊 3 次。",
                "non_ex": (
                    "Excited vocal shouting during outdoor playground play;"
                    " tapping cheeks rhythmically during music group time."
                ),
                "non_ex_zh": (
                    "在户外操场游玩时兴奋地高声欢呼；在音乐课上跟随节奏轻拍脸颊。"
                ),
                "dimensions": (
                    "Frequency: 2-4 episodes/day. Duration: 30s - 3min."
                    " Intensity: Moderate."
                ),
                "dimensions_zh": (
                    "频率：每天 2-4 次。持续时间：30秒 至 3分钟。强度：中度。"
                ),
                "triggers": (
                    "Setting Events: Overstimulating ambient noise, sudden"
                    " schedule changes.\nImmediate Triggers: Unexpected high-pitch"
                    " sounds, removal of preferred juice cup."
                ),
                "triggers_zh": (
                    "背景事件：环境噪音过度刺激、突发日程改变。\n直接触发因素：意料之外的高频噪音、偏好的果汁杯被移走。"
                ),
                "consequences": (
                    "Staff offers noise-canceling headphones and redirects to"
                    " sensory chew tool."
                ),
                "consequences_zh": "工作人员提供降噪耳机，并重新引导至口部感官咀嚼工具。",
                "qabf_summary": (
                    "Sensory/Automatic: 13/15 | Task Escape: 11/15 | Attention:"
                    " 3/15 | Tangible: 2/15 | Physical: 1/15"
                ),
                "qabf_summary_zh": (
                    "感官刺激/自动强化: 13/15 | 逃避任务: 11/15 | 社交关注: 3/15 |"
                    " 获得物质: 2/15 | 身体不适: 1/15"
                ),
                "triangulation": (
                    "Direct ABC Data (appliance noise trigger) + Indirect RBT"
                    " Interview (auditory aversion) + QABF (Sensory score"
                    " 13/15)."
                ),
                "triangulation_zh": (
                    "直接 ABC 数据（电器噪音触发） + 间接 RBT 访谈（听觉敏感过度） + QABF"
                    " 结果（感官得分 13/15）。"
                ),
                "hypothesis": (
                    "Primary Function: Escape / Regulation of Auditory"
                    " Overstimulation (Sensory Automatic Reinforcement)."
                ),
                "hypothesis_zh": (
                    "核心行为功能：逃避/调节听觉过度刺激（感官自动强化）。"
                ),
                "ferb": (
                    "Self-Regulation / Request Strategy: Point to 'Headphones'"
                    " visual symbol or independently retrieve noise-canceling"
                    " headphones from designated sensory bin."
                ),
                "ferb_zh": (
                    "功能性替代行为 (FERB)：指认“耳机”视觉标识，或自行从指定的感官箱中取出降噪耳机戴上。"
                ),
            },
            {
                "name": "3. Floor Dropping & Pica Attempts",
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
                "ex": (
                    "Dropping to playground sand and placing a handful of sand"
                    " near mouth when prompted to transition inside."
                ),
                "ex_zh": (
                    "当被提示转换进室内时，突然倒在操场沙坑里并将一把沙子拿到嘴边。"
                ),
                "non_ex": (
                    "Lying down on rest mat during scheduled nap time; placing"
                    " silicone chew toy into mouth."
                ),
                "non_ex_zh": (
                    "在计划的午休时间躺在休息垫上；将硅胶咀嚼玩具放入口中。"
                ),
                "dimensions": (
                    "Frequency: 1-3 episodes/day. Duration: 1-5min. Intensity:"
                    " Moderate."
                ),
                "dimensions_zh": (
                    "频率：每天 1-3 次。持续时间：1 至 5分钟。强度：中度。"
                ),
                "triggers": (
                    "Setting Events: Unassigned downtime, transition from"
                    " outdoor to indoor.\nImmediate Triggers: Direct"
                    " instruction to clean up preferred toys."
                ),
                "triggers_zh": (
                    "背景事件：无安排的空闲时间、从户外到室内的活动转换。\n直接触发因素：要求收拾偏好玩具的直接指令。"
                ),
                "consequences": (
                    "RBT blocks hand-to-mouth trajectory, provides oral chew"
                    " device, and delays transition demand."
                ),
                "consequences_zh": (
                    "RBT 阻挡手移向嘴部的轨迹，提供口部咀嚼器，并延缓转换要求。"
                ),
                "qabf_summary": (
                    "Task Escape: 12/15 | Tangible Access: 10/15 | Attention:"
                    " 4/15 | Sensory: 3/15 | Physical: 1/15"
                ),
                "qabf_summary_zh": (
                    "逃避任务: 12/15 | 获得物质: 10/15 | 社交关注: 4/15 | 感官刺激:"
                    " 3/15 | 身体不适: 1/15"
                ),
                "triangulation": (
                    "Direct ABC Data (playground cleanup trigger) + Indirect"
                    " Teacher Interview (transition delay) + QABF (Escape score"
                    " 12/15)."
                ),
                "triangulation_zh": (
                    "直接 ABC 数据（操场清理触发） + 间接教师访谈（转换延缓） + QABF"
                    " 结果（逃避得分 12/15）。"
                ),
                "hypothesis": (
                    "Primary Function: Transition Escape & Delay.\nSecondary"
                    " Function: Access to Outdoor Tangible Play."
                ),
                "hypothesis_zh": (
                    "核心行为功能：逃避与延缓活动转换。\n次要行为功能：继续获取户外游戏物品。"
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
        "history": (
            "Diagnosed with ASD Level 3; currently receiving 15 hrs/week of"
            " Early Intervention ABA and Speech Therapy."
        ),
        "history_zh": (
            "确诊为孤独症谱系障碍（ASD 3级）；目前每周接受 15"
            " 小时的早期干预 ABA 及言语治疗服务。"
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
                "name": "1. Task Avoidance / Elopement from Seat",
                "def": "Leaving assigned desk area without teacher permission for >5 seconds during independent academic instruction. Onset: feet leaving designated desk perimeter; Offset: returning to seat.",
                "def_zh": "在独立学业教学期间，未经教师允许离开指定的课桌区域超过5秒。以双脚离开指定课桌边界为行为开始，以返回座位为行为结束。",
                "ex": "Running out of seat to classroom carpet area during independent math worksheet.",
                "ex_zh": "在独立完成数学工作表期间，从座位上跑开并躺在教室地毯区域。",
                "non_ex": "Standing up to walk to sharpener after raising hand and receiving permission.",
                "non_ex_zh": "举手并获得许可后，站起来走到削笔器旁削铅笔。",
                "dimensions": "Frequency: 4-5 times per school day. Duration: 1-5 minutes per instance. Intensity: Low to Moderate.",
                "dimensions_zh": "频率：每个学校日 4-5 次。持续时间：每次 1-5 分钟。强度：低至中度。",
                "triggers": "Setting Events: Multi-step math tasks.\nImmediate Triggers: Presentation of 2-page math assignment.",
                "triggers_zh": "背景事件：多步骤数学任务。\n直接触发因素：发放长达2页的数学作业纸。",
                "consequences": "Staff presents 'Break' visual card; demand temporarily paused.",
                "consequences_zh": "教职工呈现“休息”视觉卡片；学业要求被暂时暂停。",
                "qabf_summary": "Task Escape: 15/15 | Attention: 5/15 | Tangible: 2/15 | Sensory: 1/15 | Physical: 0/15",
                "qabf_summary_zh": "逃避任务: 15/15 | 社交关注: 5/15 | 获得物质: 2/15 | 感官刺激: 1/15 | 身体不适: 0/15",
                "triangulation": "Direct ABC Data + Indirect IEP Interview + QABF (Escape score 15/15).",
                "triangulation_zh": "直接 ABC 数据 + 间接 IEP 访谈 + QABF 结果（逃避得分 15/15）。",
                "hypothesis": "Primary Function: Escape from Academic Demands.",
                "hypothesis_zh": "核心行为功能：逃避学业任务要求。",
                "ferb": "Hand 'Break' card to teacher or place 'Help Needed' tent on desk.",
                "ferb_zh": "向教师递交“休息”卡片，或在桌上摆放“需要帮助”提示牌。",
            }
        ],
        "strengths": "Excellent visual-spatial abilities, enthusiastic about technology and drawing.",
        "strengths_zh": "具备出色的视觉空间能力，对科技和绘画抱有极高热情。",
        "history": "Enrolled in General Education with IEP support.",
        "history_zh": "就读于普通教育班级，享有 IEP 特殊教育支持计划。",
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
                "name": "1. Vocational Task Refusal & Verbal Aggression",
                "def": "Refusing assembly or sorting demands accompanied by loud vocal threats (>75 dB) or pushing work materials away. Onset: vocal outburst or material push; Offset: 3 minutes of quiet task engagement.",
                "def_zh": "拒绝组装或分类任务，并伴随大声言语威胁（>75 dB）或推开工作材料的行为。以言语发作或推开材料为行为开始，以连续 3 分钟安静参与任务为行为结束。",
                "ex": "Shouting 'No way!', slamming assembly boxes on desk when quota is raised.",
                "ex_zh": "当提高工作配额时，大叫“绝不可能！”并将组装盒重重摔在桌上。",
                "non_ex": "Verbally requesting a 5-minute break in a normal tone.",
                "non_ex_zh": "用正常音量和语调口头提出“想要休息5分钟”。",
                "dimensions": "Frequency: 1-2 times weekly. Duration: 5-10 minutes. Intensity: Moderate.",
                "dimensions_zh": "频率：每周 1-2 次。持续时间：5-10 分钟。强度：中度。",
                "triggers": "Setting Events: Unfamiliar staff.\nImmediate Triggers: Direct instructions to complete vocational assembly quota.",
                "triggers_zh": "背景事件：不熟悉的工作人员。\n直接触发因素：要求完成职业组装配额的直接指令。",
                "consequences": "DSP offers choice board, demand temporarily paused.",
                "consequences_zh": "直属支持人员（DSP）提供选择板，任务要求被暂时暂停。",
                "qabf_summary": "Task Escape: 13/15 | Attention: 4/15 | Tangible: 3/15 | Sensory: 1/15 | Physical: 1/15",
                "qabf_summary_zh": "逃避任务: 13/15 | 社交关注: 4/15 | 获得物质: 3/15 | 感官刺激: 1/15 | 身体不适: 1/15",
                "triangulation": "Direct ABC Data + Indirect Vocational Interview + QABF (Escape score 13/15).",
                "triangulation_zh": "直接 ABC 数据 + 间接职业能力访谈 + QABF 结果（逃避得分 13/15）。",
                "hypothesis": "Primary Function: Escape from Vocational Assembly Demands.",
                "hypothesis_zh": "核心行为功能：逃避职业组装工作要求。",
                "ferb": "Verbally request '5-minute break, please' using self-advocacy phrase card.",
                "ferb_zh": "使用自我倡导短语卡口头表达：“请给我 5 分钟休息时间”。",
            }
        ],
        "strengths": "High independence in personal self-care.",
        "strengths_zh": "在个人日常生活自理方面具备极高独立性。",
        "history": "Participates in Adult Day Vocational Services under Medicaid HCBS Waiver.",
        "history_zh": "在 Medicaid HCBS 豁免计划下参与成人日间职业干预服务。",
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
  st.markdown("#### 📊 Behavior QABF Assessment Results")
  mock_qabf = generate_mock_qabf_docx(selected_cohort_key)
  st.download_button(
      label=f"📥 Download Mock QABF Results (.docx)",
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
    doc, label_en, val_en, label_trans=None, val_trans=None, is_zh=False
):
  p = doc.add_paragraph()
  p.paragraph_format.space_after = Pt(4)
  p.paragraph_format.space_before = Pt(2)

  r_lbl = p.add_run(f"{label_en}: ")
  r_lbl.bold = True
  p.add_run(f"{val_en}")

  if is_zh and label_trans and val_trans:
    p.add_run("\n")
    r_tr_lbl = p.add_run(f"[{label_trans}: ")
    r_tr_lbl.bold = True
    r_tr_lbl.italic = True
    r_tr_lbl.font.color.rgb = RGBColor(100, 100, 100)

    r_tr_val = p.add_run(f"{val_trans}]")
    r_tr_val.italic = True
    r_tr_val.font.color.rgb = RGBColor(100, 100, 100)


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
      "2. Data Sources & Assessment Methodology",
      "2. 数据来源与评估方法" if is_zh else None,
  )
  add_bi_item(
      doc,
      "Data Sources Used",
      (
          "1. Direct ABC Observations\n2. Indirect Stakeholder Interviews\n3."
          " Behavior-Specific QABF Rating Scales\n4. Environmental Baseline"
          " Analysis"
      ),
      "数据来源" if is_zh else None,
      (
          "1. 直接 ABC 观察记录\n2. 利益相关者访谈\n3. 按行为拆解的 QABF"
          " 量表评估结果\n4. 环境基线分析"
      ),
      is_zh,
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
      c_meta["strengths_zh"],
      is_zh,
  )
  add_bi_item(
      doc,
      "Clinical / Educational History",
      c_meta["history"],
      "临床/教育背景" if is_zh else None,
      c_meta["history_zh"],
      is_zh,
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
        b["def_zh"],
        is_zh,
    )
    add_bi_item(
        doc,
        "B. Examples & Non-Examples",
        f"Examples: {b['ex']}\nNon-Examples: {b['non_ex']}",
        "B. 示例与非示例" if is_zh else None,
        f"示例: {b['ex_zh']}\n非示例: {b['non_ex_zh']}",
        is_zh,
    )
    add_bi_item(
        doc,
        "C. Behavior Dimensions",
        b["dimensions"],
        "C. 行为维度 (频率/持续时间/强度)" if is_zh else None,
        b["dimensions_zh"],
        is_zh,
    )
    add_bi_item(
        doc,
        "D. Environmental Triggers & Setting Events",
        b["triggers"],
        "D. 环境触发因素与背景事件" if is_zh else None,
        b["triggers_zh"],
        is_zh,
    )
    add_bi_item(
        doc,
        "E. Maintaining Consequences",
        b["consequences"],
        "E. 维持后果与他人反应" if is_zh else None,
        b["consequences_zh"],
        is_zh,
    )
    add_bi_item(
        doc,
        "F. Behavior-Specific QABF Results",
        b["qabf_summary"],
        "F. 该行为专属 QABF 量表得分" if is_zh else None,
        b["qabf_summary_zh"],
        is_zh,
    )
    add_bi_item(
        doc,
        "G. Triangulation (Direct ABC + Indirect + QABF)",
        b["triangulation"],
        "G. 三方交叉验证 (直接数据 + 访谈 + QABF)" if is_zh else None,
        b["triangulation_zh"],
        is_zh,
    )
    add_bi_item(
        doc,
        "H. Hypothesized Function",
        b["hypothesis"],
        "H. 该行为推断功能" if is_zh else None,
        b["hypothesis_zh"],
        is_zh,
    )

  # Section 5: Overall Synthesis
  add_bi_heading(
      doc,
      1,
      "5. Synthesis & Clinical Recommendations",
      "5. 综合评估结论与临床建议" if is_zh else None,
  )
  add_bi_item(
      doc,
      "Recommendations",
      (
          "Formulate an individualized Behavior Intervention Plan (BIP)"
          " targeting each behavior's validated function through proactive"
          " modifications, Functional Communication Training (FCT), and"
          " differential reinforcement schedules."
      ),
      "干预建议" if is_zh else None,
      (
          "制定针对性的行为干预计划 (BIP)，围绕上述各行为经确证的功能，"
          "实施前因预防、功能性沟通训练 (FCT) 及差异性强化策略。"
      ),
      is_zh,
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
      "1. Student Info & Administrative Summary",
      "1. 学生/客户信息与行政摘要" if is_zh else None,
  )
  build_compact_demographics_table(doc, c_meta, is_zh)

  # Section 2: Behavior Functions & FERB Breakdown (Separated per behavior!)
  add_bi_heading(
      doc,
      1,
      "2. Target Behaviors, Functions & Replacement Skills (FERB)",
      "2. 目标行为、行为功能与替代技能 (FERB) 逐项拆解" if is_zh else None,
  )

  for idx, b in enumerate(c_meta["behaviors"], 1):
    add_bi_heading(
        doc,
        2,
        f"Target Behavior #{idx}: {b['name']}",
        f"目标行为 #{idx}: {b['name']}" if is_zh else None,
    )
    add_bi_item(
        doc,
        "Validated Function",
        b["hypothesis"],
        "确证行为功能" if is_zh else None,
        b["hypothesis_zh"],
        is_zh,
    )
    add_bi_item(
        doc,
        "Functionally Equivalent Replacement Behavior (FERB)",
        b["ferb"],
        "功能性替代行为 (FERB)" if is_zh else None,
        b["ferb_zh"],
        is_zh,
    )

  # Section 3: Proactive / Antecedent Strategies (Unified)
  add_bi_heading(
      doc,
      1,
      "3. Proactive & Antecedent Modifications (Prevention)",
      "3. 前因调整与预防策略 (统一整合)" if is_zh else None,
  )
  add_bi_item(
      doc,
      "3.1 Environmental & Priming Adaptations",
      (
          "• Provide 2-minute and 1-minute visual/auditory transition warnings"
          " prior to changing activities.\n• Offer noise-canceling headphones"
          " or move client to low-stimulation sensory area prior to task"
          " demands.\n• Break academic/vocational tasks into small, visual"
          " chunks (2-3 items per strip)."
      ),
      "3.1 环境调整与预先提示" if is_zh else None,
      (
          "• 在活动转换前提供 2分钟 及 1分钟 的视觉/听觉预先倒计时提示。\n•"
          " 在任务开始前主动提供降噪耳机或移至低刺激感官区域。\n•"
          " 将学业/职业任务拆解为小步子视觉单元（每条2-3个小任务）。"
      ),
      is_zh,
  )
  add_bi_item(
      doc,
      "3.2 Non-Contingent Reinforcement (NCR) & Choice-Making",
      (
          "• Deliver 15-30 seconds of non-contingent high quality 1:1 adult"
          " attention every 10-15 minutes during independent play.\n• Provide"
          " forced-choice options before tasks (e.g., 'Do you want red or blue"
          " crayon?')."
      ),
      "3.2 非条件性强化与选择权提供" if is_zh else None,
      (
          "• 在独立游戏期间，每 10-15 分钟主动提供 15-30"
          " 秒高质量关注（不与行为挂钩）。\n•"
          " 在任务开始前提供双选择权（例如：‘你想用红色的笔还是蓝色的笔？’）。"
      ),
      is_zh,
  )

  # Section 4: Replacement Behaviors Protocols (Unified)
  add_bi_heading(
      doc,
      1,
      "4. Replacement Behaviors & Functional Communication Training (FCT)",
      "4. 替代行为与功能性沟通训练 (FCT)" if is_zh else None,
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
      "4.1 功能性沟通协议" if is_zh else None,
      (
          "• 核心 FCT 技能：教导客户在产生烦躁情绪萌芽时，按下 AAC 沟通按键或递交 PECS"
          " 卡片表达“休息”或“帮助”。\n• 辅助渐退策略：使用由多到少（Most-to-Least）物理辅助，并在"
          " 10 天内快速渐退至手势或视觉提示。"
      ),
      is_zh,
  )
  add_bi_item(
      doc,
      "4.2 Tolerance & Delay to Reinforcement",
      (
          "• Systematically teach client to accept 'Wait 5 seconds' after"
          " requesting a break before demand is paused, gradually increasing to"
          " 30 seconds."
      ),
      "4.2 容忍度与延迟等待训练" if is_zh else None,
      (
          "• 系统性训练客户在提出“休息”请求后接受“等待 5"
          " 秒”的指令，再暂停任务，并逐步增加至等待 30 秒。"
      ),
      is_zh,
  )

  # Section 5: Reinforcement Strategies (Unified)
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
      (
          "• 在习得阶段，只要客户按下 AAC 表达“休息”，必须在 3 秒内 100%"
          " 满足其休息请求。\n•"
          " 将逃避任务与高度热情的情感口头表扬结合（如：“太棒了，你自己按按键说要休息！”）。"
      ),
      is_zh,
  )

  # Section 6: Response Strategies (Unified)
  add_bi_heading(
      doc,
      1,
      "6. Reactive Response Protocols & Extinction",
      "6. 目标行为回应与消退策略" if is_zh else None,
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
      "6.1 消退与中立物理阻挡" if is_zh else None,
      (
          "•"
          " 逃避消退：在问题行为发生时，保持平静中立表情，避免眼神接触，不进行长篇大论训诫。\n•"
          " 物理阻挡：若出现自伤或打打行为，使用软垫迅速柔和阻挡，确保安全的同时不给予额外的言语或情感反馈。"
      ),
      is_zh,
  )

  # Section 7: Safety & Data
  add_bi_heading(
      doc,
      1,
      "7. Crisis Safety & Treatment Fidelity",
      "7. 危机安全预案与执行忠实度" if is_zh else None,
  )
  add_bi_item(
      doc,
      "Data Collection & Fidelity Protocols",
      (
          "• RBTs will record daily frequency/duration of target behaviors and"
          " independent FCT requests.\n• BCBA will conduct weekly treatment"
          " fidelity observations using a 10-point checklist."
      ),
      "数据收集与忠实度核查" if is_zh else None,
      (
          "• RBT 每日记录目标行为的发生频率/持续时间及 FCT 独立使用次数。\n• BCBA"
          " 每周使用 10 项标准核查表进行 1:1 干预忠实度评估。"
      ),
      is_zh,
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
