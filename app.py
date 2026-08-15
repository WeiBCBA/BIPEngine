import io
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor


# ---------------------------------------------------------------------------
# 1. 辅助排版函数 (支持多语言/双语输出)
# ---------------------------------------------------------------------------
def add_bi_heading(doc, level, text_en, text_other=None):
    heading = doc.add_heading(level=level)
    run_en = heading.add_run(text_en)
    run_en.font.name = "Calibri"
    run_en.bold = True
    run_en.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    if text_other:
        run_other = heading.add_run(f" [{text_other}]")
        run_other.font.name = "Microsoft YaHei"
        run_other.font.size = Pt(12)
        run_other.bold = False
        run_other.font.color.rgb = RGBColor(0x59, 0x59, 0x59)


def add_bi_item(
    doc, title_en, text_en, title_other=None, text_other=None, lang_mode="en"
):
    p = doc.add_paragraph()

    # 主语言/英文标题
    run_t_en = p.add_run(f"• {title_en}: ")
    run_t_en.bold = True
    run_t_en.font.name = "Calibri"
    run_t_en.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    run_v_en = p.add_run(str(text_en))
    run_v_en.font.name = "Calibri"

    # 双语对照文本 (中文/西班牙文)
    if lang_mode != "en" and text_other:
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.left_indent = Pt(14)

        run_t_ot = p_sub.add_run(f"[{title_other or title_en}: ")
        run_t_ot.bold = True
        run_t_ot.font.name = "Microsoft YaHei"
        run_t_ot.font.size = Pt(9.5)
        run_t_ot.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

        run_v_ot = p_sub.add_run(f"{text_other}]")
        run_v_ot.font.name = "Microsoft YaHei"
        run_v_ot.font.size = Pt(9.5)
        run_v_ot.font.color.rgb = RGBColor(0x59, 0x59, 0x59)


# ---------------------------------------------------------------------------
# 2. Word 文档核心生成逻辑 (去重与优化)
# ---------------------------------------------------------------------------
def generate_exact_fba_doc(data, lang_mode="en"):
    doc = Document()

    # 文档标题
    title_p = doc.add_paragraph()
    run_t = title_p.add_run(
        "FUNCTIONAL BEHAVIOR ASSESSMENT (FBA) & BEHAVIOR INTERVENTION PLAN"
        " (BIP)"
    )
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(16)
    run_t.bold = True
    run_t.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

    if lang_mode == "zh":
        run_zh = title_p.add_run(
            "\n功能性行为评估 (FBA) 与行为干预计划 (BIP) 综合报告"
        )
        run_zh.font.name = "Microsoft YaHei"
        run_zh.font.size = Pt(13)
        run_zh.bold = True
        run_zh.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    elif lang_mode == "es":
        run_es = title_p.add_run(
            "\nEVALUACIÓN DE CONDUCTA FUNCIONAL (FBA) Y PLAN DE INTERVENCIÓN DE"
            " CONDUCTA (BIP)"
        )
        run_es.font.name = "Calibri"
        run_es.font.size = Pt(13)
        run_es.bold = True
        run_es.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

    doc.add_paragraph()

    # Section 1: Demographics
    add_bi_heading(
        doc,
        1,
        "1. Demographics & Administrative Context",
        (
            "1. 个人基本信息与行政背景"
            if lang_mode == "zh"
            else "1. Datos demográficos y contexto administrativo"
            if lang_mode == "es"
            else None
        ),
    )

    student_name = data.get("student_name", "N/A")
    dob = data.get("dob", "N/A")
    grade = data.get("grade", "N/A")
    setting = data.get("setting", "N/A")
    assessor = data.get("assessor", "N/A")
    date_str = data.get("date", "N/A")

    demo_text = (
        f"Student Name: {student_name}\n"
        f"Date of Birth: {dob}\n"
        f"Grade/Placement: {grade}\n"
        f"Primary Setting: {setting}\n"
        f"Lead Assessor/BCBA: {assessor}\n"
        f"Assessment Date: {date_str}"
    )

    demo_zh = (
        f"学生姓名: {student_name}\n"
        f"出生日期: {dob}\n"
        f"年级/班级: {grade}\n"
        f"主要环境: {setting}\n"
        f"主评估师/BCBA: {assessor}\n"
        f"评估日期: {date_str}"
        if lang_mode == "zh"
        else None
    )

    demo_es = (
        f"Nombre del estudiante: {student_name}\n"
        f"Fecha de nacimiento: {dob}\n"
        f"Grado/Ubicación: {grade}\n"
        f"Entorno principal: {setting}\n"
        f"Evaluador principal/BCBA: {assessor}\n"
        f"Fecha de evaluación: {date_str}"
        if lang_mode == "es"
        else None
    )

    add_bi_item(
        doc,
        "Client Identification Profile",
        demo_text,
        (
            "客户身份识别档案"
            if lang_mode == "zh"
            else "Perfil de identificación"
            if lang_mode == "es"
            else None
        ),
        demo_zh if lang_mode == "zh" else demo_es,
        lang_mode,
    )

    # Section 2: Data Sources & Triangulation (已优化去重)
    add_bi_heading(
        doc,
        1,
        "2. Data Sources & Triangulation Methodology",
        (
            "2. 数据来源与三方交叉验证评估算法"
            if lang_mode == "zh"
            else "2. Fuentes de datos y metodología de triangulación"
            if lang_mode == "es"
            else None
        ),
    )

    add_bi_item(
        doc,
        "Triangulation Weighting & Synthesis Algorithm",
        (
            "Data synthesis and functional hypotheses are formulated using a"
            " weighted triangulation algorithm across three primary assessment"
            " sources:\n• Direct ABC Continuous Observations: 65%\n• Indirect"
            " Stakeholder Interviews: 25%\n• Psychometric QABF Rating Scales:"
            " 10%"
        ),
        (
            "三方交叉验证算法与权重配置"
            if lang_mode == "zh"
            else "Algoritmo de triangulación y ponderación"
            if lang_mode == "es"
            else None
        ),
        (
            "本报告数据综合结论与功能推断基于以下三方评估来源的加权交叉验证算法得出：\n•"
            " 直接 ABC 连续观察数据：65%\n• 利益相关者间接访谈：25%\n• 心理测量"
            " QABF 量表评估：10%"
            if lang_mode == "zh"
            else "La síntesis de datos y las hipótesis funcionales se formulan"
            " mediante un algoritmo de triangulación ponderado:\n• Observaciones"
            " continuas ABC directas: 65%\n• Entrevistas a partes interesadas:"
            " 25%\n• Escalas psicométricas QABF: 10%"
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )

    # Section 3: Target Behaviors & Operational Definitions
    add_bi_heading(
        doc,
        1,
        "3. Target Behaviors & Operational Definitions",
        (
            "3. 目标行为与可操作性定义"
            if lang_mode == "zh"
            else "3. Conductas objetivo y definiciones operacionales"
            if lang_mode == "es"
            else None
        ),
    )

    behaviors = data.get("behaviors", [])
    for idx, b in enumerate(behaviors, 1):
        b_name = b.get("name", "Behavior")
        b_def = b.get("definition", "N/A")
        b_func = b.get("function", "N/A")

        en_desc = (
            f"Operational Definition: {b_def}\nPrimary Hypothesized Function:"
            f" {b_func}"
        )
        zh_desc = (
            f"可操作性定义: {b_def}\n核心假设功能: {b_func}"
            if lang_mode == "zh"
            else None
        )
        es_desc = (
            f"Definición operacional: {b_def}\nFunción principal hipotetizada:"
            f" {b_func}"
            if lang_mode == "es"
            else None
        )

        add_bi_item(
            doc,
            f"Target Behavior {idx}: {b_name}",
            en_desc,
            (
                f"目标行为 {idx}: {b_name}"
                if lang_mode == "zh"
                else f"Conducta objetivo {idx}: {b_name}"
                if lang_mode == "es"
                else None
            ),
            zh_desc if lang_mode == "zh" else es_desc,
            lang_mode,
        )

    # Section 4: Functional Analysis & Baseline Summary
    add_bi_heading(
        doc,
        1,
        "4. Functional Analysis & Baseline Summary",
        (
            "4. 功能分析与基线总结"
            if lang_mode == "zh"
            else "4. Análisis funcional y resumen de línea base"
            if lang_mode == "es"
            else None
        ),
    )

    fa_summary = data.get(
        "fa_summary", "Detailed baseline data available upon request."
    )
    add_bi_item(
        doc,
        "Baseline Data Synthesis",
        fa_summary,
        (
            "基线数据综合分析"
            if lang_mode == "zh"
            else "Síntesis de datos de línea base"
            if lang_mode == "es"
            else None
        ),
        (
            f"基线总结分析: {fa_summary}"
            if lang_mode == "zh"
            else f"Resumen de línea base: {fa_summary}"
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )

    # Section 5: Behavior Intervention Plan (BIP) Strategies
    add_bi_heading(
        doc,
        1,
        "5. Behavior Intervention Plan (BIP) Multi-Tiered Strategies",
        (
            "5. 行为干预计划 (BIP) 多层级策略"
            if lang_mode == "zh"
            else "5. Plan de intervención de conducta (BIP) Estrategias por niveles"
            if lang_mode == "es"
            else None
        ),
    )

    antecedents = data.get("antecedents", "N/A")
    replacements = data.get("replacements", "N/A")
    consequences = data.get("consequences", "N/A")

    add_bi_item(
        doc,
        "Proactive / Antecedent Strategies",
        antecedents,
        (
            "前因预防策略"
            if lang_mode == "zh"
            else "Estrategias de antecedente / proactivas"
            if lang_mode == "es"
            else None
        ),
        (
            f"预防策略配置: {antecedents}"
            if lang_mode == "zh"
            else f"Estrategias proactivas: {antecedents}"
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )

    add_bi_item(
        doc,
        "Teaching Replacement Behaviors (FERB)",
        replacements,
        (
            "功能替代行为教学 (FERB)"
            if lang_mode == "zh"
            else "Enseñanza de conductas de reemplazo (FERB)"
            if lang_mode == "es"
            else None
        ),
        (
            f"替代行为教学: {replacements}"
            if lang_mode == "zh"
            else f"Conductas de reemplazo: {replacements}"
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )

    add_bi_item(
        doc,
        "Reactive / Consequence Management",
        consequences,
        (
            "后果反应与强化管理"
            if lang_mode == "zh"
            else "Gestión reactiva / de consecuencias"
            if lang_mode == "es"
            else None
        ),
        (
            f"后果管理策略: {consequences}"
            if lang_mode == "zh"
            else f"Gestión de consecuencias: {consequences}"
            if lang_mode == "es"
            else None
        ),
        lang_mode,
    )

    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream


# ---------------------------------------------------------------------------
# 3. Streamlit 界面逻辑与预设数据字典 (已修正字符标点错误)
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Clinical BIP Generator", page_icon="🧩", layout="wide"
)

# 初始化预设数据模版
preset_g2 = {
    "student_name": "Alexander Chen",
    "dob": "2018-04-12",
    "grade": "Grade 2 (General Education with Learning Support)",
    "setting": "Primary Elementary Classroom & Playground",
    "assessor": "Wei, BCBA",
    "date": "2026-05-15",
    "protocol_sentences": [
        (
            "• Aligned with IDEA IEP requirements and PBIS Multi-Tiered Support"
            " Systems."
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
    ],  # 标点符号已修正为英文半角逗号
    "behaviors": [
        {
            "name": "Task Avoidance / Non-Compliance",
            "definition": (
                "Refusing to initiate or complete independent written"
                " assignments within 2 minutes of teacher instruction, manifested"
                " by pushing materials away or putting head on desk."
            ),
            "function": (
                "Escape/Avoidance of non-preferred academic demands (specifically"
                " writing tasks)."
            ),
        },
        {
            "name": "Verbal Disruption",
            "definition": (
                "Vocalizations exceeding conversational volume during quiet work"
                " time, including calling out without raising hand or making"
                " off-task comments."
            ),
            "function": "Attention-seeking from peers and adult staff.",
        },
    ],
    "fa_summary": (
        "Baseline assessment indicates non-compliance occurs in 75% of writing"
        " tasks with average latency of 4.5 minutes. Verbal disruptions occur at"
        " a rate of 3.2 instances per hour during independent work."
    ),
    "antecedents": (
        "Provide visual schedule, break tasks into 10-minute chunks using visual"
        " timer, provide choice of writing instruments."
    ),
    "replacements": (
        "Teach functional communication request: 'I need a 2-minute break'"
        " using a visual Break Card."
    ),
    "consequences": (
        "Differential Reinforcement of Alternative Behavior (DRA): Provide"
        " immediate praise and token upon appropriate break request or task"
        " initiation."
    ),
}

# 页面标题 UI
st.title("🧩 Clinical BIP / FBA Generator")
st.caption("Professional Behavior Analytic Assessment & Intervention Builder")

# 多语言模式选择器
lang_mode = st.radio(
    "Select Output Language / 选择输出语言:",
    options=["en", "zh", "es"],
    format_func=lambda x: {
        "en": "English Only",
        "zh": "English + 中文对照",
        "es": "English + Español",
    }[x],
    horizontal=True,
)

st.divider()

# 表单配置区
col1, col2 = st.columns(2)
with col1:
    student_name = st.text_input(
        "Student Name", value=preset_g2["student_name"]
    )
    dob = st.text_input("Date of Birth", value=preset_g2["dob"])
    grade = st.text_input("Grade/Placement", value=preset_g2["grade"])

with col2:
    setting = st.text_input("Primary Setting", value=preset_g2["setting"])
    assessor = st.text_input("Lead Assessor/BCBA", value=preset_g2["assessor"])
    date_str = st.text_input("Assessment Date", value=preset_g2["date"])

st.subheader("BIP Strategies Configuration")
antecedents = st.text_area(
    "Proactive / Antecedent Strategies", value=preset_g2["antecedents"]
)
replacements = st.text_area(
    "Replacement Behaviors (FERB)", value=preset_g2["replacements"]
)
consequences = st.text_area(
    "Reactive / Consequence Management", value=preset_g2["consequences"]
)

# 封装数据结构提交生成
input_data = {
    "student_name": student_name,
    "dob": dob,
    "grade": grade,
    "setting": setting,
    "assessor": assessor,
    "date": date_str,
    "behaviors": preset_g2["behaviors"],
    "fa_summary": preset_g2["fa_summary"],
    "antecedents": antecedents,
    "replacements": replacements,
    "consequences": consequences,
}

st.divider()

if st.button("🚀 Generate FBA/BIP Document (.docx)", type="primary"):
    doc_bytes = generate_exact_fba_doc(input_data, lang_mode=lang_mode)
    st.success("Document generated successfully!")
    st.download_button(
        label="📥 Download Word File",
        data=doc_bytes,
        file_name=f"FBA_BIP_{student_name.replace(' ', '_')}.docx",
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
    )
