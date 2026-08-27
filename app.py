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
    add_styled_paragraph(doc, f"    - Elementi: {dati.get('conn_trave_pilastro_elementi', 'N.D.')} — Peso: {dati.get('conn_trave_pilastro_kg', 'N.D.')}")
    add_styled_paragraph(doc, f"• Base Pilastro / Fondazione: {dati.get('conn_pilastro_fondazione_tipo', 'N.D.')}")
    add_styled_paragraph(doc, f"    - Elementi: {dati.get('conn_pilastro_fondazione_elementi', 'N.D.')} — Peso: {dati.get('conn_pilastro_fondazione_kg', 'N.D.')}")

    add_section_title("7. Protezione Antincendio")
    add_styled_paragraph(doc, f"• Classe di Resistenza al Fuoco: {dati.get('classe_resistenza_fuoco', 'R 60')}")
    add_styled_paragraph(doc, f"• Superficie Acciaio da Trattare: {dati.get('mq_intumescente', 'N.D.')}")
    add_styled_paragraph(doc, f"• Ciclo Antincendio: {dati.get('dettaglio_verniciatura', 'N.D.')}")

    add_section_title("8. Note Tecniche e Considerazioni Finali")
    add_styled_paragraph(doc, dati.get('note_tecniche', 'Nessuna nota aggiuntiva.'))

    p_footer = doc.add_paragraph()
    p_footer.paragraph_format.space_before = Pt(24)
    r_foot = p_footer.add_run("Documento generato automaticamente tramite generatore interno WolfSystem | Uso tecnico e commerciale preliminare.")
    r_foot.font.size = Pt(8)
    r_foot.font.color.rgb = RGBColor(120, 120, 120)

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# --- CONFIGURAZIONE INTERFACCIA ---
st.set_page_config(page_title="Predimensionamento Strutturale IA", layout="wide")
st.title("Generatore Offerte Tecniche e Dimensionamento IA 🏗️")

with st.sidebar:
    st.header("Impostazioni IA")
    api_key = st.text_input("Inserisci qui la tua API Key di Google", type="password")
    st.info("L'API Key serve per far leggere il capitolato all'Intelligenza Artificiale.")
    st.markdown("---")
    st.markdown("**WolfSystem / Tecnico:** Generazione report su carta intestata, calcoli Neve, Vento, Sisma (NTC 2018) e nodi.")

st.subheader("Analisi Capitolato / Appunti di Progetto e File CAD")

file_cad_caricato = st.file_uploader("📂 Carica un file CAD (.dxf) con le specifiche o note del progetto (facoltativo)", type=["dxf"])

testo_da_cad = ""
if file_cad_caricato is not None:
    bytes_data = file_cad_caricato.read()
    stream = io.BytesIO(bytes_data)
    doc_dxf = ezdxf.read(stream)
    msp = doc_dxf.modelspace()
    testi_estratto = []
    for entity in msp:
        if entity.dxftype() == 'TEXT':
            testi_estratto.append(entity.dxf.text)
        elif entity.dxftype() == 'MTEXT':
            testi_estratto.append(entity.text)
    testo_da_cad = "\n".join(testi_estratto)
    st.success(f"File CAD '{file_cad_caricato.name}' letto con successo!")

testo_commerciale = st.text_area(
    "Incolla qui le note del progetto o il capitolato:", 
    height=150,
    value="",
    placeholder="Incolla qui le note del progetto, i dati dimensionali, il luogo di costruzione o le richieste specifiche del capitolato..."
)

if st.button("Esegui Dimensionamento Completo (Neve, Vento, Sisma)", type="primary"):
    if not api_key:
        st.error("Inserisci prima l'API Key nella barra laterale!")
    else:
        testo_totale_analisi = testo_commerciale + "\n\n--- NOTE DAL CAD ---\n" + testo_da_cad
        if not testo_totale_analisi.strip():
            st.warning("Inserisci del testo nel riquadro o carica un file CAD prima di eseguire il dimensionamento.")
        else:
            genai.configure(api_key=api_key)
            try:
                model = genai.GenerativeModel(
                    model_name='gemini-3.6-flash',
                    generation_config={"response_mime_type": "application/json"}
                )
                
                prompt = f"""
                Sei un ingegnere strutturista senior in WolfSystem, esperto in prefabbricazione, NTC 2018 (analisi combinata Neve, Vento e Sisma) e nodi esecutivi.
                Analizza il testo tecnico fornito e calcola un predimensionamento strutturale conservativo. In particolare, deduci rigorosamente:
                1. I parametri neve (qsk) e vento (Zona Vento, pressione).
                2. I parametri sismici secondo NTC 2018 in base alla località (Zona Sismica, accelerazione di picco ag, Classe d'Uso della costruzione e Fattore di struttura q).
                
                Restituisci ESATTAMENTE e unicamente un oggetto JSON valido (senza blocchi markdown) con queste chiavi esatte:
                - "luogo": stringa
                - "qsk": float
                - "zona_vento": stringa
                - "pressione_vento": stringa
                - "zona_sismica": stringa (es. "Zona 2 (ag = 0.225g)")
                - "classe_uso": stringa (es. "Classe II - Edifici industriali/commerciali")
                - "fattore_struttura_q": stringa (es. "q = 2.0 (strutture intelaiate)")
                - "interasse_portali": float
                - "interasse_arcarecci": float
                - "sezione_arcarecci": stringa
                - "verifica_arcarecci": stringa
                - "travi_legno": stringa
                - "travi_acciaio": stringa
                - "travi_cap": stringa
                - "pilastri_legno": stringa
                - "pilastri_acciaio": stringa
                - "pilastri_cap": stringa
                - "controventi_copertura_legno": stringa
                - "controventi_copertura_acciaio": stringa
                - "controventi_copertura_pos": stringa
                - "controventi_parete_legno": stringa
                - "controventi_parete_acciaio": stringa
                - "controventi_parete_pos": stringa
                - "conn_trave_pilastro_tipo": stringa
                - "conn_trave_pilastro_elementi": stringa
                - "conn_trave_pilastro_kg": stringa
                - "conn_pilastro_fondazione_tipo": stringa
                - "conn_pilastro_fondazione_elementi": stringa
                - "conn_pilastro_fondazione_kg": stringa
                - "classe_resistenza_fuoco": stringa
                - "mq_intumescente": stringa
                - "dettaglio_verniciatura": stringa
                - "note_tecniche": stringa
                
                Testo:
                "{testo_totale_analisi}"
                """
                
                with st.spinner('Elaborazione calcoli strutturali (Neve, Vento, Sisma NTC 2018)...'):
                    risposta_ia = model.generate_content(prompt)
                    testo_risposta = risposta_ia.text.strip()
                    if testo_risposta.startswith("```json"):
                        testo_risposta = testo_risposta[7:]
                    if testo_risposta.endswith("```"):
                        testo_risposta = testo_risposta[:-3]
                    
                    dati = json.loads(testo_risposta.strip())
                    st.session_state['dati_ultimi'] = dati
                    st.success("Dimensionamento completo con verifica Sismica completato!")
                    
            except Exception as e:
                st.error(f"Errore: {e}")

if 'dati_ultimi' in st.session_state:
    dati = st.session_state['dati_ultimi']
    st.markdown("---")
    
    col_dl1, col_dl2, col_dl3 = st.columns([1, 2, 1])
    with col_dl2:
        word_file = genera_word_report(dati)
        st.download_button(
            label="📄 Scarica Relazione Tecnica su Carta Intestata WolfSystem (.docx)",
            data=word_file,
            file_name=f"Relazione_WolfSystem_{dati.get('luogo', 'Progetto').replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
    
    st.markdown("---")
    
    # 1. Dati Climatici e Sismici
    st.markdown("### 📍 1. Dati geometrici, climatici e sismici (NTC 2018)")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Località", dati.get("luogo", "Bolzano"))
    c2.metric("Carico Neve (qsk)", f"{dati.get('qsk', 1.5)} kN/m²")
    c3.metric("Zona Vento", dati.get("zona_vento", "Zona 3"))
    c4.metric("Zona Sismica", dati.get("zona_sismica", "Zona 2"))
    
    c5, c6, c7 = st.columns(3)
    c5.metric("Pressione Vento", dati.get("pressione_vento", "0.80 kN/m²"))
    c6.metric("Classe d'Uso", dati.get("classe_uso", "Classe II"))
    c7.metric("Fattore di Struttura", dati.get("fattore_struttura_q", "q = 2.0"))
    
    st.markdown("---")
    
    # 2. Arcarecci
    st.markdown("### 🪵 2. Arcarecci di Copertura")
    st.info(f"**Sezione Consigliata:** {dati.get('sezione_arcarecci', 'N.D.')} | **Stato:** {dati.get('verifica_arcarecci', 'Verificato')}")
    
    st.markdown("---")
    
    # 3. Travi
    st.markdown("### 📐 3. Travi Principali / Portali (Confronto Tecnologico)")
    col_t1, col_t2, col_t3 = st.columns(3)
    with col_t1:
        st.markdown("#### 🌲 Legno Lamellare")
        st.success(dati.get('travi_legno', 'N.D.'))
    with col_t2:
        st.markdown("#### ⚙️ Acciaio")
        st.warning(dati.get('travi_acciaio', 'N.D.'))
    with col_t3:
        st.markdown("#### 🏛️ C.a.p.")
        st.error(dati.get('travi_cap', 'N.D.'))
    
    st.markdown("---")
    
    # 4. Pilastri
    st.markdown("### 🏛️ 4. Pilastri (Confronto Tecnologico)")
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1:
        st.markdown("#### 🌲 Legno Lamellare")
        st.success(dati.get('pilastri_legno', 'N.D.'))
    with col_p2:
        st.markdown("#### ⚙️ Acciaio")
        st.warning(dati.get('pilastri_acciaio', 'N.D.'))
    with col_p3:
        st.markdown("#### 🏛️ C.a.p.")
        st.error(dati.get('pilastri_cap', 'N.D.'))
    
    st.markdown("---")
    
    # 5. Controventi
    st.markdown("### 🔗 5. Stabilizzazione e Controventi (Azioni Orizzontali)")
    col_cv1, col_cv2 = st.columns(2)
    with col_cv1:
        st.markdown("#### 🛡️ Controventi di Copertura (Falda)")
        st.write(f"📍 **Posizionamento:** {dati.get('controventi_copertura_pos', 'N.D.')}")
        st.info(f"🌲 **Opzione Legno:** {dati.get('controventi_copertura_legno', 'N.D.')}")
        st.info(f"⚙️ **Opzione Acciaio:** {dati.get('controventi_copertura_acciaio', 'N.D.')}")
    with col_cv2:
        st.markdown("#### 🧱 Controventi di Parete (Baraccatura)")
        st.write(f"📍 **Posizionamento:** {dati.get('controventi_parete_pos', 'N.D.')}")
        st.warning(f"🌲 **Opzione Legno:** {dati.get('controventi_parete_legno', 'N.D.'))
        st.warning(f"⚙️ **Opzione Acciaio:** {dati.get('controventi_parete_acciaio', 'N.D.')}")
    
    st.markdown("---")
    
    # 6. Connessioni
    st.markdown("### 🔩 6. Dimensionamento Dettagliato Connessioni e Nodi")
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        st.markdown("#### 🔗 Connessione Pilastro / Trave di Copertura")
        st.info(f"**Tipologia Nodo:** {dati.get('conn_trave_pilastro_tipo', 'N.D.')}")
        st.write(f"- **Organi di Collegamento:** {dati.get('conn_trave_pilastro_elementi', 'N.D.')}")
        st.metric("Peso Acciaio Connessione", dati.get('conn_trave_pilastro_kg', 'N.D.'))
    with col_n2:
        st.markdown("#### ⚓ Connessione Pilastro / Fondazione")
        st.warning(f"**Tipologia Base:** {dati.get('conn_pilastro_fondazione_tipo', 'N.D.')}")
        st.write(f"- **Organi di Ancoraggio:** {dati.get('conn_pilastro_fondazione_elementi', 'N.D.')}")
        st.metric("Peso Acciaio Ancoraggi/Piastra", dati.get('conn_pilastro_fondazione_kg', 'N.D.'))
    
    st.markdown("---")
    
    # 7. Antincendio
    st.markdown("### 🔥 7. Requisiti di Resistenza al Fuoco e Vernice Intumescente")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        st.metric("Classe di Resistenza Richiesta", dati.get('classe_resistenza_fuoco', 'R 60'))
    with col_f2:
        st.metric("Superficie Acciaio da Trattare", dati.get('mq_intumescente', 'Non specificato'))
    st.info(f"**Specifiche Ciclo Antincendio:** {dati.get('dettaglio_verniciatura', 'N.D.')}")
    
    st.markdown("---")
    
    # 8. Note
    st.markdown("### 📝 8. Relazione e Note Tecniche")
    st.write(dati.get("note_tecniche", "Nessuna nota aggiuntiva."))
