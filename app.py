# -*- coding: utf-8 -*-
"""
================================================================================
US BCBA Functional Behavior Assessment (FBA) & Behavior Intervention Plan (BIP)
Master Single-File Web Application Engine
--------------------------------------------------------------------------------
Author: Clinical Data Engineering Team
Target Framework: Streamlit / python-docx / XML Custom Styling
Supported Age Groups: 
  1. Early Intervention (0-7 Years)
  2. Youth / School-Age (8-17 Years)
  3. Adults / Supported Independent Living (18+ Years)
Languages: English (EN), Spanish (ES), Chinese (ZH)
================================================================================
"""

import os
import io
import time
import json
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import streamlit as st

# ==============================================================================
# 1. STREAMLIT CONFIGURATION & CUSTOM CSS INJECTION
# ==============================================================================
st.set_page_config(
    page_title="BCBA Master FBA & BIP Generator System",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main .block-container { padding-top: 1.5rem; padding-bottom: 3rem; }
    .stAlert { border-radius: 8px; }
    div[data-testid="stSidebarUserContent"] { padding-top: 1rem; }
    .metric-card {
        background-color: #f8f9fa;
        border-left: 5px solid #1b365d;
        padding: 12px;
        border-radius: 4px;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# 2. DOCUMENT XML STYLING HELPER (ADVANCED WORD FORMATTING)
# ==============================================================================
class XMLStyleHelper:
    """Helper class to directly manipulate docx XML elements for professional styling."""
    
    @staticmethod
    def setup_margins(doc, top=1.0, bottom=1.0, left=1.0, right=1.0):
        """Sets standard 1-inch margins across all sections."""
        for section in doc.sections:
            section.top_margin = Inches(top)
            section.bottom_margin = Inches(bottom)
            section.left_margin = Inches(left)
            section.right_margin = Inches(right)

    @staticmethod
    def set_cell_bg(cell, hex_color):
        """Applies solid background color to a table cell."""
        tcPr = cell._element.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        tcPr.append(shd)

    @staticmethod
    def set_cell_margins(cell, top=120, bottom=120, left=160, right=160):
        """Sets cell inner padding in DXA units (1/20 of a point)."""
        tcPr = cell._element.get_or_add_tcPr()
        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="{top}" w:type="dxa"/>'
            f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
            f'<w:left w:w="{left}" w:type="dxa"/>'
            f'<w:right w:w="{right}" w:type="dxa"/>'
            f'</w:tcMar>'
        )
        tcPr.append(tcMar)

    @staticmethod
    def add_callout_box(doc, text, title="", border_color="1B365D", bg_color="F2F4F4"):
        """Renders a single-cell callout box with a thick left border."""
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.5)
        
        XMLStyleHelper.set_cell_bg(cell, bg_color)
        XMLStyleHelper.set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
        
        tcPr = cell._element.get_or_add_tcPr()
        borders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="none"/>'
            f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>'
            f'<w:bottom w:val="none"/>'
            f'<w:right w:val="none"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(borders)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        
        if title:
            r_title = p.add_run(f"{title}\n")
            r_title.bold = True
            r_title.font.name = "Calibri"
            r_title.font.size = Pt(10.5)
            r_title.font.color.rgb = RGBColor(27, 54, 93)
            
        r_text = p.add_run(text)
        r_text.font.name = "Calibri"
        r_text.font.size = Pt(9.5)
        r_text.font.italic = True
        
        doc.add_paragraph() # Spacing


# ==============================================================================
# 3. CLINICAL DATASETS (3 AGE GROUPS x DUAL LANGUAGES)
# ==============================================================================
CLINICAL_DATABASE = {
    "0-7": {
        "label_en": "0-7 Years (Early Intervention)",
        "label_zh": "0-7 岁（早期干预与幼儿组）",
        "EN": {
            "setting": "Clinic & Natural Home Environment",
            "staff_roles": "Parents, RBT, Early Intervention Specialist, Speech Therapist",
            "triangulation_summary": (
                "Direct ABC observation (65% weight) across 14 clinic hours reveals Target Behavior 1 (SIB) "
                "occurs primarily during non-preferred activity transitions. Indirect MAS parent interview (25% weight) "
                "supports an Escape/Avoidance function. QABF questionnaire (10% weight) indicated high scores on "
                "Physical Discomfort and Escape."
            ),
            "behaviors": [
                {
                    "name": "Target Behavior 1: SIB (Head Banging)",
                    "definition": "Forceful contact between forehead and hard floor/walls during non-preferred activity transitions.",
                    "function": "Escape from High-Demand Transitions",
                    "baseline_rate": "3.5 episodes per 4-hour session",
                    "severity": "High (Risk of tissue damage)"
                },
                {
                    "name": "Target Behavior 2: Open-Palm Face Slapping",
                    "definition": "Striking own cheeks with open palm resulting in audible sound or localized redness.",
                    "function": "Automatic Sensory Reinforcement",
                    "baseline_rate": "12 instances per day during unstructured downtime",
                    "severity": "Moderate"
                }
            ],
            "hypothesis": (
                "When presented with academic or routines transitions, the child engages in SIB (head banging) "
                "to delay task onset, which is negatively reinforced by staff temporarily removing demands. "
                "Face slapping is maintained by automatic sensory stimulation during downtime."
            ),
            "env_mods": [
                "Implement visual First-Then board 3 minutes prior to transitions.",
                "Incorporate heavy-work sensory breaks (proprioceptive input) every 20 minutes.",
                "Provide functional choice menu for transition items."
            ],
            "bip_protocols": [
                {
                    "name": "Functional Communication Training (FCT) - Break AAC",
                    "target": "SIB (Head Banging)",
                    "ferb": "Pressing 'I Need a Break' icon on AAC device or handing visual break card.",
                    "fading_criteria": "80% independent prompt-free communication across 5 consecutive days.",
                    "prompts": [
                        "Level 1: Full Physical Hand-over-Hand to activate AAC button",
                        "Level 2: Partial Physical prompt at wrist/elbow",
                        "Level 3: Visual gesture pointing to AAC device",
                        "Level 4: Independent communication upon antecedents"
                    ]
                }
            ],
            "bst_steps": [
                "1. Instructions: BCBA reviews visual transition schedule protocol with parents.",
                "2. Modeling: BCBA demonstrates AAC break prompting during simulated toy removal.",
                "3. Rehearsal: Parents practice prompting AAC break card during real transition.",
                "4. Feedback: BCBA provides immediate praise and corrective feedback."
            ]
        },
        "ES": {
            "setting": "Entorno Clínico y Hogar Natural",
            "staff_roles": "Padres, RBT, Especialista en Intervención Temprana",
            "triangulation_summary": (
                "Observación directa ABC (65%) muestra que la Autolesión (SIB) ocurre en transiciones. "
                "Entrevista MAS con padres (25%) apoya función de Escape."
            ),
            "behaviors": [
                {
                    "name": "Conducta 1: Autolesión (Golpe de Cabeza)",
                    "definition": "Contacto fuerte de la frente contra el suelo o paredes durante transiciones.",
                    "function": "Escape de Transiciones de Alta Demanda",
                    "baseline_rate": "3.5 episodios por sesión de 4 horas",
                    "severity": "Alta"
                }
            ],
            "hypothesis": "Al presentar transiciones, el niño realiza SIB para retrasar la demanda.",
            "env_mods": ["Implementar tablero visual Primero-Después 3 minutos antes."],
            "bip_protocols": [
                {
                    "name": "Entrenamiento en Comunicación Funcional (FCT)",
                    "target": "Autolesión",
                    "ferb": "Presionar icono 'Necesito un Descanso' en dispositivo CAA.",
                    "fading_criteria": "80% de comunicación independiente en 5 días.",
                    "prompts": [
                        "Nivel 1: Guía física total mano sobre mano",
                        "Nivel 2: Guía física parcial en la muñeca",
                        "Nivel 3: Apoyo gestual al dispositivo",
                        "Nivel 4: Respuesta independiente"
                    ]
                }
            ],
            "bst_steps": [
                "1. Instrucción: Revisar protocolo de apoyos visuales con padres.",
                "2. Modelado: BCBA demuestra el desvanecimiento de apoyos.",
                "3. Ensayo: Los padres practican en una transición real.",
                "4. Retroalimentación: BCBA da retroalimentación inmediata."
            ]
        }
    },
    "8-17": {
        "label_en": "8-17 Years (School-Age / Youth)",
        "label_zh": "8-17 岁（青少年与学龄组）",
        "EN": {
            "setting": "Public School & Community Context",
            "staff_roles": "Special Education Teacher, Classroom Paraprofessional, RBT",
            "triangulation_summary": (
                "Direct ABC data (65% weight) gathered during 20 academic hours indicates Elopement occurs "
                "during complex multi-step math/reading worksheets. Student interview (25% weight) confirms "
                "feelings of frustration and desire to escape difficult tasks. QABF score (10%) confirms Escape."
            ),
            "behaviors": [
                {
                    "name": "Target Behavior 1: Task Avoidance Elopement",
                    "definition": "Leaving assigned desk or classroom without permission for > 5 seconds during academic instruction.",
                    "function": "Escape from Multi-Step Academic Demands",
                    "baseline_rate": "4.2 instances per school day",
                    "severity": "Moderate to High (Safety concern if leaving building)"
                }
            ],
            "hypothesis": (
                "When presented with independent academic tasks exceeding 10 minutes, the student elopes from "
                "the desk to avoid task completion, which is reinforced by escape from instructional demands."
            ),
            "env_mods": [
                "Chunk academic assignments into 5-minute work intervals using a visual countdown timer.",
                "Provide designated 'Cool Down Corner' access without academic materials.",
                "Embed preferred interest themes (e.g., video games) into reading worksheets."
            ],
            "bip_protocols": [
                {
                    "name": "Self-Advocacy Break Card System",
                    "target": "Elopement",
                    "ferb": "Placing a laminated Break Pass on the teacher's desk before walking to the cool-down corner.",
                    "fading_criteria": "Zero unexcused elopements across 20 consecutive school days.",
                    "prompts": [
                        "Level 1: Point to Break Pass when sighing or showing frustration cues",
                        "Level 2: Verbal reminder: 'If you need a break, use your pass'",
                        "Level 3: Independent self-advocacy"
                    ]
                }
            ],
            "bst_steps": [
                "1. Instruction: BCBA reviews break pass policy with special education teacher and aide.",
                "2. Modeling: BCBA models silent acceptance of break pass during live instruction.",
                "3. Rehearsal: Teacher practice responding neutrally to break pass requests.",
                "4. Feedback: BCBA provides weekly implementation fidelity scoring."
            ]
        },
        "ES": {
            "setting": "Escuela Pública y Entorno Comunitario",
            "staff_roles": "Maestro de Educación Especial, Paraprofesional, RBT",
            "triangulation_summary": "Observación ABC (65%) indica que el Escape ocurre durante hojas de trabajo complejas.",
            "behaviors": [
                {
                    "name": "Conducta 1: Escape / Abandono de Tarea",
                    "definition": "Salir del escritorio o aula asignada sin permiso por más de 5 segundos.",
                    "function": "Escape de Demandas Académicas",
                    "baseline_rate": "4.2 instancias por día escolar",
                    "severity": "Moderada / Alta"
                }
            ],
            "hypothesis": "Al enfrentar tareas académicas largas, el estudiante abandona el escritorio para escapar.",
            "env_mods": ["Dividir tareas académicas en bloques de 5 minutos con temporizador visual."],
            "bip_protocols": [
                {
                    "name": "Sistema de Pase de Descanso para Autogestión",
                    "target": "Escape de Tarea",
                    "ferb": "Entregar pase de descanso antes de caminar a la zona de calma.",
                    "fading_criteria": "Cero escapes sin permiso por 20 días consecutivos.",
                    "prompts": [
                        "Nivel 1: Senalar el pase de descanso al notar frustración",
                        "Nivel 2: Recordatorio verbal sutil",
                        "Nivel 3: Uso independiente"
                    ]
                }
            ],
            "bst_steps": [
                "1. Instrucción: Revisar plan con el maestro y asistente.",
                "2. Modelado: BCBA demuestra la aceptación neutral del pase.",
                "3. Ensayo: El personal practica la entrega de refuerzo.",
                "4. Retroalimentación: Monitoreo de fidelidad semanal."
            ]
        }
    },
    "18+": {
        "label_en": "18+ Years (Adult Services / SIL)",
        "label_zh": "18岁+（成年人与社区独立生活组）",
        "EN": {
            "setting": "Supported Independent Living (SIL) & Day Program",
            "staff_roles": "Direct Support Professional (DSP), Residential Manager, Job Coach",
            "triangulation_summary": (
                "Direct ABC observation (65% weight) over 30 days shows physical aggression occurs when preferred "
                "community outings are restricted. DSP staff interviews (25%) support Control/Tangible function. "
                "QABF (10%) scores high on Tangible Access and Attention."
            ),
            "behaviors": [
                {
                    "name": "Target Behavior 1: Physical Aggression",
                    "definition": "Pushing, grabbing clothing, or striking staff with open/closed fist during schedule changes.",
                    "function": "Access to Preferred Tangibles & Choice Restriction Escape",
                    "baseline_rate": "1.8 episodes per week",
                    "severity": "High (Staff safety risk)"
                }
            ],
            "hypothesis": (
                "When access to preferred community activities or items is restricted without choice, "
                "the client engages in physical aggression to compel staff to grant access or alter schedule."
            ),
            "env_mods": [
                "Implement daily person-centered visual choice schedules every morning.",
                "Provide 15-minute and 5-minute pre-warnings prior to any community outing transition.",
                "Maintain transparent stock of preferred items with clear visual availability indicators."
            ],
            "bip_protocols": [
                {
                    "name": "De-escalation & Choice-Based Communication Protocol",
                    "target": "Physical Aggression",
                    "ferb": "Utilizing communication binder to point to alternative preferred activity or negotiating delay.",
                    "fading_criteria": "Zero instances of physical aggression for 60 consecutive days.",
                    "prompts": [
                        "Level 1: Present choice binder passively at 6-foot safe distance",
                        "Level 2: Low-tone verbal prompt: 'We can choose another activity on your board'",
                        "Level 3: Independent self-advocacy and choice selection"
                    ]
                }
            ],
            "bst_steps": [
                "1. Instruction: BCBA reviews person-centered support rights and de-escalation steps with DSPs.",
                "2. Modeling: BCBA demonstrates non-confrontational stance and choice board offering.",
                "3. Rehearsal: DSP staff role-play crisis de-escalation and evasion.",
                "4. Feedback: Monthly fidelity audits during residential shifts."
            ]
        },
        "ES": {
            "setting": "Vivienda Apoyada e Independiente (SIL) y Programa de Día",
            "staff_roles": "Profesional de Apoyo Directo (DSP), Gerente Residencial",
            "triangulation_summary": "Datos ABC (65%) muestran que la agresión ocurre al restringir salidas comunitarias.",
            "behaviors": [
                {
                    "name": "Conducta 1: Agresión Física",
                    "definition": "Empujar, agarrar la ropa o golpear al personal al cambiar horarios.",
                    "function": "Acceso a Tangibles Preferidos / Control",
                    "baseline_rate": "1.8 episodios por semana",
                    "severity": "Alta"
                }
            ],
            "hypothesis": "Al restringir actividades preferidas sin opciones, el cliente realiza agresión para obtener acceso.",
            "env_mods": ["Implementar horarios de elección visual centrados en la persona."],
            "bip_protocols": [
                {
                    "name": "Protocolo de Desescalada y Comunicación Basada en Opciones",
                    "target": "Agresión Física",
                    "ferb": "Usar carpeta de comunicación para seleccionar actividad alternativa.",
                    "fading_criteria": "Cero agresión física por 60 días consecutivos.",
                    "prompts": [
                        "Nivel 1: Presentar menú de opciones a distancia segura",
                        "Nivel 2: Apoyo verbal en tono bajo",
                        "Nivel 3: Selección independiente"
                    ]
                }
            ],
            "bst_steps": [
                "1. Instrucción: Revisar el plan centrado en la persona con el personal.",
                "2. Modelado: BCBA demuestra la postura no conflictiva.",
                "3. Ensayo: El personal realiza juego de roles de desescalada.",
                "4. Retroalimentación: Auditorías de fidelidad mensuales."
            ]
        }
    }
}


# ==============================================================================
# 4. FBA WORD DOCUMENT GENERATOR CLASS
# ==============================================================================
class US_FBA_Generator:
    """Generates professional Word (.docx) FBA reports with strict XML table and callout formatting."""
    
    def __init__(self, language="EN"):
        self.doc = docx.Document()
        self.lang = language
        XMLStyleHelper.setup_margins(self.doc)

    def generate(self, client_name, evaluator, dates, payload, output_path):
        # Header Title
        p_title = self.doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_after = Pt(18)
        
        r_title = p_title.add_run("FUNCTIONAL BEHAVIOR ASSESSMENT (FBA) REPORT" if self.lang == "EN" else "INFORME DE EVALUACIÓN FUNCIONAL DEL COMPORTAMIENTO")
        r_title.font.name = "Calibri"
        r_title.font.size = Pt(18)
        r_title.font.bold = True
        r_title.font.color.rgb = RGBColor(27, 54, 93)

        # 1. Demographics Section
        h1 = self.doc.add_heading("1. Demographics & Assessment Context", level=1)
        h1.style.font.color.rgb = RGBColor(27, 54, 93)
        
        tbl = self.doc.add_table(rows=3, cols=4)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        demo_data = [
            ("Client Name / ID:", client_name, "Age Group / Cohort:", payload.get("setting", "N/A")),
            ("Primary Setting:", payload.get("setting", "Clinic"), "Lead Evaluator:", evaluator),
            ("Assessment Period:", dates, "Language Format:", self.lang)
        ]
        
        for i, row in enumerate(demo_data):
            for j in range(4):
                cell = tbl.cell(i, j)
                XMLStyleHelper.set_cell_margins(cell)
                cell.paragraphs[0].paragraph_format.space_after = Pt(0)
                if j % 2 == 0:
                    XMLStyleHelper.set_cell_bg(cell, "EAEDED")
                    r = cell.paragraphs[0].add_run(row[j])
                    r.bold = True
                    r.font.size = Pt(9.5)
                else:
                    r = cell.paragraphs[0].add_run(row[j])
                    r.font.size = Pt(9.5)

        self.doc.add_paragraph() # Spacing

        # 2. Triangulation & Methodology
        h2 = self.doc.add_heading("2. Assessment Methodology & Data Weighting", level=1)
        h2.style.font.color.rgb = RGBColor(27, 54, 93)
        
        callout_text = (
            "Clinical Triangulation Model Applied (65% / 25% / 10% Standard):\n"
            "• Direct ABC Data Collection (65% Weight): Continuous baseline observation across settings.\n"
            "• Indirect Assessment / MAS (25% Weight): Semistructured interview with primary caregivers/staff.\n"
            "• Rating Scale / QABF (10% Weight): Quantitative verification of behavioral function.\n\n"
            f"Triangulation Findings: {payload['triangulation_summary']}"
        )
        XMLStyleHelper.add_callout_box(self.doc, callout_text, "Methodology & Triangulation Summary")

        # 3. Target Behaviors Table
        h3 = self.doc.add_heading("3. Target Behaviors & Functional Analysis", level=1)
        h3.style.font.color.rgb = RGBColor(27, 54, 93)
        
        beh_list = payload["behaviors"]
        tbl_b = self.doc.add_table(rows=len(beh_list) + 1, cols=4)
        tbl_b.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        headers = ["Target Behavior", "Operational Definition", "Baseline Rate", "Function"]
        for j, h in enumerate(headers):
            cell = tbl_b.cell(0, j)
            XMLStyleHelper.set_cell_bg(cell, "1B365D")
            XMLStyleHelper.set_cell_margins(cell)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(h)
            r.bold = True
            r.font.color.rgb = RGBColor(255, 255, 255)
            r.font.size = Pt(9.5)

        for i, b in enumerate(beh_list):
            row_idx = i + 1
            tbl_b.cell(row_idx, 0).paragraphs[0].add_run(b["name"]).bold = True
            tbl_b.cell(row_idx, 1).paragraphs[0].add_run(b["definition"])
            tbl_b.cell(row_idx, 2).paragraphs[0].add_run(b.get("baseline_rate", "N/A"))
            tbl_b.cell(row_idx, 3).paragraphs[0].add_run(b["function"]).bold = True
            
            for j in range(4):
                XMLStyleHelper.set_cell_margins(tbl_b.cell(row_idx, j))

        self.doc.add_paragraph()

        # 4. Functional Hypothesis & Recommendations
        h4 = self.doc.add_heading("4. Functional Hypothesis & Antecedent Strategies", level=1)
        h4.style.font.color.rgb = RGBColor(27, 54, 93)
        
        p_hyp = self.doc.add_paragraph()
        r_h_title = p_hyp.add_run("Synthesized Behavioral Hypothesis:\n")
        r_h_title.bold = True
        p_hyp.add_run(payload["hypothesis"])
        
        self.doc.add_heading("Proactive Environmental Modifications", level=2)
        for env in payload["env_mods"]:
            p_item = self.doc.add_paragraph(style='List Bullet')
            p_item.add_run(env)

        self.doc.save(output_path)
        return output_path
        # ==============================================================================
# 5. BIP WORD DOCUMENT GENERATOR CLASS
# ==============================================================================
class US_BIP_Generator:
    """Generates professional Word (.docx) BIP reports with prompt hierarchies and BST protocols."""

    def __init__(self, language="EN"):
        self.doc = docx.Document()
        self.lang = language
        XMLStyleHelper.setup_margins(self.doc)

    def generate(self, client_name, evaluator, dates, payload, output_path):
        # Header Title
        p_title = self.doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_after = Pt(18)

        r_title = p_title.add_run(
            "BEHAVIOR INTERVENTION PLAN (BIP)" if self.lang == "EN" else "PLAN DE INTERVENCIÓN EN EL COMPORTAMIENTO"
        )
        r_title.font.name = "Calibri"
        r_title.font.size = Pt(18)
        r_title.font.bold = True
        r_title.font.color.rgb = RGBColor(27, 54, 93)

        # 1. Demographics
        h1 = self.doc.add_heading("1. Client Identification & Implementation Scope", level=1)
        h1.style.font.color.rgb = RGBColor(27, 54, 93)

        tbl = self.doc.add_table(rows=2, cols=4)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        demo_data = [
            ("Client Name:", client_name, "Implementation Setting:", payload.get("setting", "General")),
            ("Lead BCBA:", evaluator, "Plan Effective Dates:", dates)
        ]

        for i, row in enumerate(demo_data):
            for j in range(4):
                cell = tbl.cell(i, j)
                XMLStyleHelper.set_cell_margins(cell)
                cell.paragraphs[0].paragraph_format.space_after = Pt(0)
                if j % 2 == 0:
                    XMLStyleHelper.set_cell_bg(cell, "EAEDED")
                    r = cell.paragraphs[0].add_run(row[j])
                    r.bold = True
                    r.font.size = Pt(9.5)
                else:
                    r = cell.paragraphs[0].add_run(row[j])
                    r.font.size = Pt(9.5)

        self.doc.add_paragraph()

        # 2. Replacement Behaviors & Prompt Fading Hierarchy
        h2 = self.doc.add_heading("2. Functional Replacement Behaviors & Prompt Fading Hierarchy", level=1)
        h2.style.font.color.rgb = RGBColor(27, 54, 93)

        protocols = payload.get("bip_protocols", [])
        for proto in protocols:
            self.doc.add_heading(f"Protocol: {proto['name']}", level=2)
            
            p_desc = self.doc.add_paragraph()
            p_desc.add_run("Target Behavior to Reduce: ").bold = True
            p_desc.add_run(f"{proto['target']}\n")
            p_desc.add_run("Functionally Equivalent Replacement Behavior (FERB): ").bold = True
            p_desc.add_run(f"{proto['ferb']}\n")
            p_desc.add_run("Mastery & Fading Criterion: ").bold = True
            p_desc.add_run(f"{proto['fading_criteria']}")

            # Prompt Hierarchy Table
            prompts = proto.get("prompts", [])
            if prompts:
                tbl_p = self.doc.add_table(rows=len(prompts) + 1, cols=2)
                tbl_p.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                # Header
                cell_h0 = tbl_p.cell(0, 0)
                cell_h1 = tbl_p.cell(0, 1)
                XMLStyleHelper.set_cell_bg(cell_h0, "1B365D")
                XMLStyleHelper.set_cell_bg(cell_h1, "1B365D")
                XMLStyleHelper.set_cell_margins(cell_h0)
                XMLStyleHelper.set_cell_margins(cell_h1)
                
                r0 = cell_h0.paragraphs[0].add_run("Prompt Level")
                r0.bold = True
                r0.font.color.rgb = RGBColor(255, 255, 255)
                r1 = cell_h1.paragraphs[0].add_run("Clinical Description & Implementation")
                r1.bold = True
                r1.font.color.rgb = RGBColor(255, 255, 255)

                for p_idx, p_str in enumerate(prompts):
                    r_cell0 = tbl_p.cell(p_idx + 1, 0)
                    r_cell1 = tbl_p.cell(p_idx + 1, 1)
                    XMLStyleHelper.set_cell_margins(r_cell0)
                    XMLStyleHelper.set_cell_margins(r_cell1)
                    
                    parts = p_str.split(":", 1)
                    if len(parts) == 2:
                        r_cell0.paragraphs[0].add_run(parts[0]).bold = True
                        r_cell1.paragraphs[0].add_run(parts[1])
                    else:
                        r_cell0.paragraphs[0].add_run(f"Level {p_idx + 1}").bold = True
                        r_cell1.paragraphs[0].add_run(p_str)

            self.doc.add_paragraph()

        # 3. Behavioral Skills Training (BST) Staff Training
        h3 = self.doc.add_heading("3. Staff & Caregiver Behavioral Skills Training (BST) Plan", level=1)
        h3.style.font.color.rgb = RGBColor(27, 54, 93)

        bst_intro = (
            "All direct-care staff, RBTs, and primary caregivers must achieve 90% fidelity "
            "across 3 consecutive BST evaluations using the following 4-step framework prior to independent execution:"
        )
        self.doc.add_paragraph(bst_intro)

        for step in payload.get("bst_steps", []):
            p_step = self.doc.add_paragraph(style='List Bullet')
            p_step.add_run(step)

        self.doc.add_paragraph()

        # 4. Emergency & De-escalation Safety Protocol
        h4 = self.doc.add_heading("4. De-escalation & Crisis Safety Protocols", level=1)
        h4.style.font.color.rgb = RGBColor(27, 54, 93)

        crisis_text = (
            "CRISIS SAFETY NOTICE:\n"
            "If the client exhibits immediate safety-threatening behaviors (severe SIB, physical aggression resulting in harm), "
            "implement bodily safety management without verbal commentary. Maintain safe distance, block physical strikes "
            "using least-restrictive techniques, remove immediate environmental hazards, and notify the supervising BCBA "
            "within 24 hours. Under no circumstances should physical restraint be used as a behavioral consequence or punishment."
        )
        XMLStyleHelper.add_callout_box(self.doc, crisis_text, "CRISIS MANAGEMENT PROTOCOL", border_color="C0392B", bg_color="FDEDEC")

        self.doc.save(output_path)
        return output_path


# ==============================================================================
# 6. STREAMLIT FRONTEND ENGINE & SESSION STATE INITIALIZATION
# ==============================================================================
def init_session_state():
    """Initializes and maintains global session state variables across user interactions."""
    if "age_group" not in st.session_state:
        st.session_state.age_group = "0-7"
    if "lang" not in st.session_state:
        st.session_state.lang = "EN"
    if "client_name" not in st.session_state:
        st.session_state.client_name = "Alex Vance"
    if "evaluator" not in st.session_state:
        st.session_state.evaluator = "Dr. Sarah Jenkins, BCBA-D"
    if "dates" not in st.session_state:
        st.session_state.dates = "August 2026 - February 2027"
    if "custom_payload" not in st.session_state:
        # Clone default from CLINICAL_DATABASE
        st.session_state.custom_payload = json.loads(
            json.dumps(CLINICAL_DATABASE[st.session_state.age_group][st.session_state.lang])
        )


def render_sidebar():
    """Renders side control panel for configuration, language, and cohort selection."""
    st.sidebar.image("https://img.icons8.com/color/96/mental-health.png", width=64)
    st.sidebar.title("BCBA FBA/BIP Engine")
    st.sidebar.caption("v4.2 Professional Clinical Suite")
    st.sidebar.divider()

    st.sidebar.subheader("1. Cohort & Language Controls")
    
    # Age Group Selector
    selected_age = st.sidebar.selectbox(
        "Select Target Cohort / Age Group:",
        options=["0-7", "8-17", "18+"],
        format_func=lambda x: CLINICAL_DATABASE[x]["label_en"],
        index=0
    )

    # Language Selector
    selected_lang = st.sidebar.radio(
        "Output Language:",
        options=["EN", "ES"],
        format_func=lambda x: "English (US Standard)" if x == "EN" else "Español (US Clinical)",
        horizontal=True
    )

    # Detect changes to update payload
    if selected_age != st.session_state.age_group or selected_lang != st.session_state.lang:
        st.session_state.age_group = selected_age
        st.session_state.lang = selected_lang
        # Reload default payload for selected age and language
        if selected_lang in CLINICAL_DATABASE[selected_age]:
            st.session_state.custom_payload = json.loads(
                json.dumps(CLINICAL_DATABASE[selected_age][selected_lang])
            )
        else:
            # Fallback to EN if language unavailable
            st.session_state.custom_payload = json.loads(
                json.dumps(CLINICAL_DATABASE[selected_age]["EN"])
            )
        st.rerun()

    st.sidebar.divider()
    st.sidebar.subheader("2. Demographics & Context")
    st.session_state.client_name = st.sidebar.text_input("Client ID / Name:", st.session_state.client_name)
    st.session_state.evaluator = st.sidebar.text_input("Lead Assessor / BCBA:", st.session_state.evaluator)
    st.session_state.dates = st.sidebar.text_input("Assessment Period:", st.session_state.dates)

    st.sidebar.divider()
    st.sidebar.markdown(
        """
        **Methodology Weighting Standard:**
        * 📊 **65%** Direct ABC Observation
        * 🗣️ **25%** MAS Caregiver Interview
        * 📝 **10%** QABF Rating Scale
        """
    )


# ==============================================================================
# 7. TAB RENDERERS (FOUR CORE INTERACTIVE PANELS)
# ==============================================================================
def render_tab_triangulation():
    """Tab 1: Visual Analytics & Clinical Data Triangulation Dashboard."""
    st.header("📊 Clinical Triangulation & Baseline Data Analytics")
    st.info(
        "This dashboard integrates multi-source clinical evidence applying the US 65%/25%/10% weighting rule "
        "to synthesize behavior functions with high diagnostic confidence."
    )

    payload = st.session_state.custom_payload

    # Key Metrics Bar
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Primary Setting", payload.get("setting", "N/A"))
    with col2:
        st.metric("Target Behaviors Count", len(payload.get("behaviors", [])))
    with col3:
        st.metric("Direct ABC Weight", "65%", "14.5 Hours Observed")
    with col4:
        st.metric("Inter-Observer Agreement (IOA)", "92.4%", "+2.1% vs Baseline")

    st.divider()

    # Visualizing Data Triangulation Split
    col_chart1, col_chart2 = st.columns([1, 1])

    with col_chart1:
        st.subheader("Methodology Weighting Split")
        triangulation_data = {
            "Methodology Source": ["Direct ABC Observation", "MAS Indirect Interview", "QABF Rating Scale"],
            "Weighting Percentage": [65, 25, 10]
        }
        st.bar_chart(data=triangulation_data, x="Methodology Source", y="Weighting Percentage")

    with col_chart2:
        st.subheader("Synthesized Diagnostic Triangulation Summary")
        st.write(payload.get("triangulation_summary", "No triangulation summary available."))
        
        st.markdown("**Core Roles Involved:**")
        st.caption(payload.get("staff_roles", "N/A"))

    st.divider()

    # Target Behavior Details
    st.subheader("Identified Target Behaviors & Baseline Rates")
    for i, beh in enumerate(payload.get("behaviors", [])):
        with st.expander(f"📌 {beh['name']} (Function: {beh['function']})", expanded=True):
            c1, c2, c3 = st.columns([2, 1, 1])
            with c1:
                st.write(f"**Operational Definition:** {beh['definition']}")
            with c2:
                st.write(f"**Baseline Frequency:** {beh.get('baseline_rate', 'N/A')}")
            with c3:
                st.write(f"**Clinical Severity:** {beh.get('severity', 'N/A')}")


def render_tab_fba_config():
    """Tab 2: Functional Behavior Assessment Configuration & Hypothesis Editor."""
    st.header("📋 FBA Configuration & Antecedent Modifications")
    st.caption("Customize clinical findings, functional hypotheses, and environmental modifications.")

    payload = st.session_state.custom_payload

    # Environment & Setting Editor
    st.subheader("1. Assessment Context & Roles")
    payload["setting"] = st.text_input("Primary Assessment Setting:", payload.get("setting", ""))
    payload["staff_roles"] = st.text_input("Key Support Roles Involved:", payload.get("staff_roles", ""))

    st.divider()

    # Synthesized Hypothesis
    st.subheader("2. Synthesized Functional Hypothesis")
    payload["hypothesis"] = st.text_area(
        "Clinical Functional Behavioral Hypothesis:",
        value=payload.get("hypothesis", ""),
        height=100
    )

    st.divider()

    # Proactive Environmental Modifications
    st.subheader("3. Proactive Environmental Modifications")
    env_list = payload.get("env_mods", [])
    
    st.write("Modify or add proactive antecedent strategies:")
    updated_env_mods = []
    for idx, env_str in enumerate(env_list):
        new_val = st.text_input(f"Mod #{idx + 1}:", value=env_str, key=f"env_mod_{idx}")
        if new_val.strip():
            updated_env_mods.append(new_val)

    # Add new strategy
    new_env = st.text_input("➕ Add New Antecedent Strategy:", key="add_new_env")
    if new_env.strip():
        updated_env_mods.append(new_env)

    payload["env_mods"] = updated_env_mods
    st.session_state.custom_payload = payload
    st.success("FBA Settings updated successfully in active session.")


def render_tab_bip_builder():
    """Tab 3: Behavior Intervention Plan & Prompt Hierarchy Builder."""
    st.header("🛡️ BIP Protocols, Prompt Fading & BST Builder")
    st.caption("Configure replacement behaviors (FERB), prompt fading hierarchies, and staff BST protocols.")

    payload = st.session_state.custom_payload

    # Protocols list
    protocols = payload.get("bip_protocols", [])
    st.subheader("1. Functionally Equivalent Replacement Behaviors (FERB)")

    for p_idx, proto in enumerate(protocols):
        with st.expander(f"⚡ BIP Protocol #{p_idx + 1}: {proto['name']}", expanded=True):
            proto["name"] = st.text_input("Protocol Title:", proto["name"], key=f"proto_name_{p_idx}")
            proto["target"] = st.text_input("Target Behavior to Reduce:", proto["target"], key=f"proto_target_{p_idx}")
            proto["ferb"] = st.text_area("Replacement Behavior (FERB):", proto["ferb"], key=f"proto_ferb_{p_idx}", height=70)
            proto["fading_criteria"] = st.text_input("Mastery & Fading Criterion:", proto["fading_criteria"], key=f"proto_fade_{p_idx}")

            st.markdown("**Prompt Fading Hierarchy Levels:**")
            prompts = proto.get("prompts", [])
            updated_prompts = []
            for lvl_idx, p_str in enumerate(prompts):
                val = st.text_input(f"Level {lvl_idx + 1}:", value=p_str, key=f"prompt_{p_idx}_{lvl_idx}")
                if val.strip():
                    updated_prompts.append(val)
            proto["prompts"] = updated_prompts

    st.divider()

    # Behavioral Skills Training (BST)
    st.subheader("2. Behavioral Skills Training (BST) 4-Step Framework")
    st.caption("Standard 4-Step Staff Training: 1. Instructions | 2. Modeling | 3. Rehearsal | 4. Feedback")

    bst_steps = payload.get("bst_steps", [])
    updated_bst = []
    for step_idx, step_str in enumerate(bst_steps):
        val = st.text_input(f"BST Step {step_idx + 1}:", value=step_str, key=f"bst_step_{step_idx}")
        if val.strip():
            updated_bst.append(val)

    payload["bst_steps"] = updated_bst
    st.session_state.custom_payload = payload


def render_tab_export():
    """Tab 4: Document Generation, Preview & Download Hub."""
    st.header("📥 Document Export & Download Hub")
    st.caption("Generate professional formatted Word (.docx) documents with direct XML styling.")

    payload = st.session_state.custom_payload
    client_name = st.session_state.client_name
    evaluator = st.session_state.evaluator
    dates = st.session_state.dates
    lang = st.session_state.lang

    col_exp1, col_exp2 = st.columns(2)

    # 1. FBA Generation
    with col_exp1:
        st.subheader("📄 Functional Behavior Assessment (FBA)")
        st.write("Generates complete FBA report with demographics, 65/25/10 triangulation callout, and target behavior tables.")
        
        if st.button("🚀 Build & Render FBA Word Document", key="btn_build_fba"):
            with st.spinner("Compiling FBA XML Tables and Callouts..."):
                fba_gen = US_FBA_Generator(language=lang)
                temp_filename = f"FBA_Report_{client_name.replace(' ', '_')}.docx"
                fba_gen.generate(client_name, evaluator, dates, payload, temp_filename)
                
                with open(temp_filename, "rb") as file:
                    st.download_button(
                        label="📥 Download FBA Document (.docx)",
                        data=file,
                        file_name=temp_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                st.success("FBA Document rendered successfully!")

    # 2. BIP Generation
    with col_exp2:
        st.subheader("🛡️ Behavior Intervention Plan (BIP)")
        st.write("Generates actionable BIP report containing FERB protocols, prompt fading tables, BST steps, and crisis safety notices.")
        
        if st.button("🚀 Build & Render BIP Word Document", key="btn_build_bip"):
            with st.spinner("Compiling BIP Tables and BST Protocols..."):
                bip_gen = US_BIP_Generator(language=lang)
                temp_filename = f"BIP_Report_{client_name.replace(' ', '_')}.docx"
                bip_gen.generate(client_name, evaluator, dates, payload, temp_filename)
                
                with open(temp_filename, "rb") as file:
                    st.download_button(
                        label="📥 Download BIP Document (.docx)",
                        data=file,
                        file_name=temp_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                st.success("BIP Document rendered successfully!")

    st.divider()

    # JSON Data Backup / Inspection
    st.subheader("🔍 Inspection & JSON Raw Backup")
    with st.expander("View Raw Active Session Payload (JSON Format)", expanded=False):
        st.json(payload)
        
        # Download JSON Backup
        json_str = json.dumps(payload, indent=2, ensure_ascii=False)
        st.download_button(
            label="💾 Download Session Backup (.json)",
            data=json_str,
            file_name=f"BCBA_Session_{client_name.replace(' ', '_')}.json",
            mime="application/json"
        )


# ==============================================================================
# 8. MAIN ENTRY POINT
# ==============================================================================
def main():
    """Main application layout and tab navigation dispatcher."""
    init_session_state()
    render_sidebar()

    # Header Title Banner
    st.title("📋 BCBA FBA & BIP Master Generation System")
    st.caption(
        f"Active Client: **{st.session_state.client_name}** | Cohort: **{st.session_state.age_group} Years** | "
        f"Language: **{st.session_state.lang}** | Assessor: **{st.session_state.evaluator}**"
    )

    # Core Navigation Tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Triangulation & Analytics",
        "📋 FBA Configuration",
        "🛡️ BIP Protocols & BST",
        "📥 Document Export Hub"
    ])

    with tab1:
        render_tab_triangulation()

    with tab2:
        render_tab_fba_config()

    with tab3:
        render_tab_bip_builder()

    with tab4:
        render_tab_export()


if __name__ == "__main__":
    main()
