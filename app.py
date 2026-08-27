import streamlit as st
import google.generativeai as genai
import json
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io
import ezdxf

# --- FUNZIONE PER STILI WORD ---
def add_styled_paragraph(doc, text, style_type='Normal', bold=False, font_size=11, color_rgb=(51,51,51), space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    run.bold = bold
    run.font.name = 'Arial'
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(*color_rgb)
    return p

# --- FUNZIONE PER GENERARE IL DOCUMENTO WORD SU CARTA INTESTATA WOLFSYSTEM ---
def genera_word_report(dati):
    doc = Document()
    
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    # INTESTAZIONE AZIENDALE
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_table.autofit = False
    
    cell_left = header_table.cell(0, 0)
    cell_right = header_table.cell(0, 1)
    cell_left.width = Inches(4.0)
    cell_right.width = Inches(2.5)
    
    p_logo = cell_left.paragraphs[0]
    p_logo.paragraph_format.space_after = Pt(0)
    r_logo = p_logo.add_run("WOLFSYSTEM")
    r_logo.bold = True
    r_logo.font.name = 'Arial'
    r_logo.font.size = Pt(18)
    r_logo.font.color.rgb = RGBColor(0, 0, 0)
    
    p_sub = cell_left.add_paragraph()
    p_sub.paragraph_format.space_after = Pt(0)
    r_sub = p_sub.add_run("Divisione Strutture Prefabbricate | Ufficio Tecnico & Estimatori")
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(9)
    r_sub.font.color.rgb = RGBColor(100, 100, 100)
    
    p_meta = cell_right.paragraphs[0]
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_meta.paragraph_format.space_after = Pt(0)
    r_meta1 = p_meta.add_run(f"Località: {dati.get('luogo', 'Bolzano')}\n")
    r_meta1.font.size = Pt(9)
    r_meta1.font.color.rgb = RGBColor(80, 80, 80)
    r_meta2 = p_meta.add_run("Data: Rilevamento Automatico IA")
    r_meta2.font.size = Pt(8)
    r_meta2.font.color.rgb = RGBColor(120, 120, 120)
    
    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_after = Pt(12)
    
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(12)
    h1.paragraph_format.space_after = Pt(6)
    r_h1 = h1.add_run("RELAZIONE TECNICA DI PREDIMENSIONAMENTO STRUTTURALE (NTC 2018)")
    r_h1.font.name = 'Arial'
    r_h1.font.size = Pt(14)
    r_h1.font.color.rgb = RGBColor(0, 0, 0)
    
    def add_section_title(title_text):
        h = doc.add_heading(level=2)
        h.paragraph_format.space_before = Pt(10)
        h.paragraph_format.space_after = Pt(4)
        r = h.add_run(title_text)
        r.font.name = 'Arial'
        r.font.size = Pt(12)
        r.font.color.rgb = RGBColor(50, 50, 50)
        return h

    add_section_title("1. Parametri Geometrici, Climatici e Sismici (NTC 2018)")
    add_styled_paragraph(doc, f"• Località di intervento: {dati.get('luogo', 'Bolzano')}")
    add_styled_paragraph(doc, f"• Carico neve al suolo (qsk): {dati.get('qsk', 1.5)} kN/m²")
    add_styled_paragraph(doc, f"• Azione del Vento: {dati.get('zona_vento', 'Zona 3')} — Pressione: {dati.get('pressione_vento', '0.80 kN/m²')}")
    add_styled_paragraph(doc, f"• Azione Sismica: {dati.get('zona_sismica', 'Zona 2')} — Classe d'Uso: {dati.get('classe_uso', 'Classe II')} — Fattore q: {dati.get('fattore_struttura_q', 'q = 2.0')}")
    add_styled_paragraph(doc, f"• Interasse Portali: {dati.get('interasse_portali', 6.0)} m | Interasse Arcarecci: {dati.get('interasse_arcarecci', 1.5)} m")

    add_section_title("2. Arcarecci di Copertura")
    add_styled_paragraph(doc, f"• Sezione Consigliata: {dati.get('sezione_arcarecci', 'N.D.')}")
    add_styled_paragraph(doc, f"• Stato Verifiche: {dati.get('verifica_arcarecci', 'Verificato')}")

    add_section_title("3. Travi Principali / Portali (Confronto Tecnologico)")
    add_styled_paragraph(doc, f"• Legno Lamellare: {dati.get('travi_legno', 'N.D.')}")
    add_styled_paragraph(doc, f"• Acciaio: {dati.get('travi_acciaio', 'N.D.')}")
    add_styled_paragraph(doc, f"• C.a.p.: {dati.get('travi_cap', 'N.D.')}")

    add_section_title("4. Pilastri (Confronto Tecnologico)")
    add_styled_paragraph(doc, f"• Legno Lamellare: {dati.get('pilastri_legno', 'N.D.')}")
    add_styled_paragraph(doc, f"• Acciaio: {dati.get('pilastri_acciaio', 'N.D.')}")
    add_styled_paragraph(doc, f"• C.a.p.: {dati.get('pilastri_cap', 'N.D.')}")

    add_section_title("5. Stabilizzazione e Controventi (Azioni Orizzontali Vento/Sisma)")
    add_styled_paragraph(doc, f"• Copertura - Posizione: {dati.get('controventi_copertura_pos', 'N.D.')}")
    add_styled_paragraph(doc, f"    - Opzione Legno: {dati.get('controventi_copertura_legno', 'N.D.')}")
    add_styled_paragraph(doc, f"    - Opzione Acciaio: {dati.get('controventi_copertura_acciaio', 'N.D.')}")
    add_styled_paragraph(doc, f"• Parete - Posizione: {dati.get('controventi_parete_pos', 'N.D.')}")
    add_styled_paragraph(doc, f"    - Opzione Legno: {dati.get('controventi_parete_legno', 'N.D.')}")
    add_styled_paragraph(doc, f"    - Opzione Acciaio: {dati.get('controventi_parete_acciaio', 'N.D.')}")

    add_section_title("6. Dettaglio Connessioni e Nodi Strutturali")
    add_styled_paragraph(doc, f"• Nodo Pilastro / Trave: {dati.get('conn_trave_pilastro_tipo', 'N.D.')}")
    add_styled_paragraph(doc, f"    - Elementi: {dati.get('conn_trave_pilastro_elementi', 'N.D.')} — Peso: {dati.get('conn_trave_
