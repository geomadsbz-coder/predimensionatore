import streamlit as st
import google.generativeai as genai
import json
from docx import Document
import io

# --- FUNZIONE PER GENERARE IL DOCUMENTO WORD ---
def genera_word_report(dati):
    doc = Document()
    doc.add_heading('Relazione Tecnica di Predimensionamento', 0)
    
    doc.add_heading('1. Dati Geometrici e Climatici (NTC 2018)', level=1)
    doc.add_paragraph(f"Località: {dati.get('luogo', 'Bolzano')}")
    doc.add_paragraph(f"Carico Neve Base (qsk): {dati.get('qsk', 1.5)} kN/m²")
    doc.add_paragraph(f"Zona Vento: {dati.get('zona_vento', 'N.D.')}")
    doc.add_paragraph(f"Pressione Vento: {dati.get('pressione_vento', 'N.D.')}")
    doc.add_paragraph(f"Interasse Portali: {dati.get('interasse_portali', 6.0)} m")
    doc.add_paragraph(f"Interasse Arcarecci: {dati.get('interasse_arcarecci', 1.5)} m")
    
    doc.add_heading('2. Arcarecci di Copertura', level=1)
    doc.add_paragraph(f"Sezione Consigliata: {dati.get('sezione_arcarecci', 'N.D.')}")
    doc.add_paragraph(f"Stato Verifiche: {dati.get('verifica_arcarecci', 'Verificato')}")
    
    doc.add_heading('3. Travi Principali / Portali (Confronto Tecnologico)', level=1)
    doc.add_paragraph(f"Legno Lamellare: {dati.get('travi_legno', 'N.D.')}")
    doc.add_paragraph(f"Acciaio: {dati.get('travi_acciaio', 'N.D.')}")
    doc.add_paragraph(f"C.a.p.: {dati.get('travi_cap', 'N.D.')}")
    
    doc.add_heading('4. Pilastri (Confronto Tecnologico)', level=1)
    doc.add_paragraph(f"Legno Lamellare: {dati.get('pilastri_legno', 'N.D.')}")
    doc.add_paragraph(f"Acciaio: {dati.get('pilastri_acciaio', 'N.D.')}")
    doc.add_paragraph(f"C.a.p.: {dati.get('pilastri_cap', 'N.D.')}")
    
    doc.add_heading('5. Stabilizzazione e Controventi', level=1)
    doc.add_paragraph(f"Posizionamento Controventi Copertura: {dati.get('controventi_copertura_pos', 'N.D.')}")
    doc.add_paragraph(f"  - Opzione Legno: {dati.get('controventi_copertura_legno', 'N.D.')}")
    doc.add_paragraph(f"  - Opzione Acciaio: {dati.get('controventi_copertura_acciaio', 'N.D.')}")
    doc.add_paragraph(f"Posizionamento Controventi Parete: {dati.get('controventi_parete_pos', 'N.D.')}")
    doc.add_paragraph(f"  - Opzione Legno: {dati.get('controventi_parete_legno', 'N.D.')}")
    doc.add_paragraph(f"  - Opzione Acciaio: {dati.get('controventi_parete_acciaio', 'N.D.')}")
    
    doc.add_heading('6. Dettaglio Connessioni e Nodi', level=1)
    doc.add_paragraph(f"Connessione Pilastro / Trave: {dati.get('conn_trave_pilastro_tipo', 'N.D.')}")
    doc.add_paragraph(f"  - Organi di collegamento: {dati.get('conn_trave_pilastro_elementi', 'N.D.')}")
    doc.add_paragraph(f"  - Peso Acciaio: {dati.get('conn_trave_pilastro_kg', 'N.D.')}")
    doc.add_paragraph(f"Connessione Pilastro / Fondazione: {dati.get('conn_pilastro_fondazione_tipo', 'N.D.')}")
    doc.add_paragraph(f"  - Organi di ancoraggio: {dati.get('conn_pilastro_fondazione_elementi', 'N.D.')}")
    doc.add_paragraph(f"  - Peso Acciaio: {dati.get('conn_pilastro_fondazione_kg', 'N.D.')}")
    
    doc.add_heading('7. Protezione Antincendio', level=1)
    doc.add_paragraph(f"Classe di Resistenza al Fuoco: {dati.get('classe_resistenza_fuoco', 'R 60')}")
    doc.add_paragraph(f"Superficie Acciaio da Trattare: {dati.get('mq_intumescente', 'N.D.')}")
    doc.add_paragraph(f"Ciclo Antincendio: {dati.get('dettaglio_verniciatura', 'N.D.')}")
    
    doc.add_heading('8. Note Tecniche e Relazione', level=1)
    doc.add_paragraph(dati.get('note_tecniche', 'Nessuna nota aggiuntiva.'))
    
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
    st.markdown("**WolfSystem / Tecnico:** Strumenti di calcolo automatico avanzato, carichi NTC 2018, nodi e report in Word.")

st.subheader("Analisi Capitolato / Appunti di Progetto")
testo_commerciale = st.text_area(
    "Incolla qui le note del progetto o il capitolato:", 
    height=150,
    value="""N° 19 file di arcarecci in legno lamellare qualità industria non impregnato posti ad interasse secondo adeguato calcolo statico, con luce statica massima di 6,00 m, collegati meccanicamente alla struttura sottostante.
N° 4 travi diagonali di testata in legno lamellare qualità industria non impregnato con luce statica massima di 9,00 m, collegate meccanicamente alla struttura sottostante.
Stabilizzazione sul piano orizzontale e verticale della struttura.
Baraccatura di parete composta da elementi in legno posti ad interasse variabile con luce statica variabile.
Tutti gli elementi strutturali in acciaio saranno protetti in modo adeguato per la verifica al fuoco.
Interasse telaio 6,00ml
Luogo di costruzione: 32044 Pieve di Cadore (BL)"""
)

if st.button("Esegui Dimensionamento Conservativo con Vento", type="primary"):
    if not api_key:
        st.error("Inserisci prima l'API Key nella barra laterale!")
    elif not testo_commerciale:
        st.warning("Incolla del testo da analizzare.")
    else:
        genai.configure(api_key=api_key)
        try:
            model = genai.GenerativeModel(
                model_name='gemini-3.6-flash',
                generation_config={"response_mime_type": "application/json"}
            )
            
            prompt = f"""
            Sei un ingegnere strutturista senior esperto in edifici prefabbricati industriali, analisi dei carichi climatici secondo NTC 2018 (Neve e Vento) e progettazione esecutiva di nodi e strutture. 
            Analizza il testo tecnico fornito e calcola un predimensionamento strutturale conservativo e robusto. 
            In particolare, in base alla località o al contesto, calcola o deduci rigorosamente i parametri del Vento (Zona Vento NTC 2018, velocità di riferimento del vento vb,0, e pressione del vento di progetto) oltre al carico neve (qsk).
            
            Restituisci ESATTAMENTE e unicamente un oggetto JSON valido (senza blocchi di codice markdown attorno) con queste chiavi esatte:
            - "luogo": stringa (località del progetto, se non specificata scrivi "Bolzano")
            - "qsk": float (carico neve al suolo in kN/m², es. 1.50)
            - "zona_vento": stringa (es. "Zona 3 (vb,0 = 27 m/s)")
            - "pressione_vento": stringa (es. "0.85 kN/m² (Pressione di picco stimata)")
            - "interasse_portali": float (in metri, se non specificato 6.0)
            - "interasse_arcarecci": float (in metri, se non specificato 1.5)
            - "sezione_arcarecci": stringa (es. "Lamellare GL24h - 140x280 mm")
            - "verifica_arcarecci": stringa (es. "Verificato a flessione, taglio, freccia e suzione vento")
            - "travi_legno": stringa (es. "Trave a falda curva GL30h - 180x1120 mm")
            - "travi_acciaio": stringa (es. "Profilo IPE 550 / Traliccio pesante in acciaio S355")
            - "travi_cap": stringa (es. "Trave a tegolo alare precompresso in C.a.p.")
            - "pilastri_legno": stringa (es. "Lamellare GL32c - 280x480 mm")
            - "pilastri_acciaio": stringa (es. "Profilo HEB 300 S355")
            - "pilastri_cap": stringa (es. "Pilastro prefabbricato in C.a.p. sezione 50x50 cm")
            - "controventi_copertura_legno": stringa (es. "Diagonali in legno lamellare GL24h - 160x200 mm")
            - "controventi_copertura_acciaio": stringa (es. "Doppie croci di Sant'Andrea in tondo d'acciaio Ø24 mm con tenditori registrabili per azioni di vento e sisma")
            - "controventi_copertura_pos": stringa (es. "N° 2 campi di falda completi in corrispondenza delle campate di testata")
            - "controventi_parete_legno": stringa (es. "Baraccatura pesante e diagonali in legno lamellare GL24h - 160x220 mm")
            - "controventi_parete_acciaio": stringa (es. "Diagonali verticali rigide in profilati cavi strutturali RHS 140x140x8 mm")
            - "controventi_parete_pos": stringa (es. "Campate perimetrali di testata e pareti longitudinali intermedie")
            - "conn_trave_pilastro_tipo": stringa (es. "Connessione a momento rigida con doppia lama d'acciaio interna a scomparsa ad alto spessore e piastra di colmo saldata")
            - "conn_trave_pilastro_elementi": stringa (es. "N° 20 spinotti in acciaio ad alta resistenza Ø16 mm disposti su maglia fitta + copripiastre esterne bullonate")
            - "conn_trave_pilastro_kg": stringa (es. "45.0 kg per nodo (configurazione pesante)")
            - "conn_pilastro_fondazione_tipo": stringa (es. "Piastra di base d'acciaio S355 spessore 35 mm irrigidita da n° 4 fazzoletti triangolari per lato e collare di base")
            - "conn_pilastro_fondazione_elementi": stringa (es. "N° 8 tirafondi di ancoraggio ad alta resistenza M30 L=1200 mm in acciaio 8.8 con dima di posa")
            - "conn_pilastro_fondazione_kg": stringa (es. "95.0 kg per plinto di fondazione")
            - "classe_resistenza_fuoco": stringa (es. "R 60")
            - "mq_intumescente": stringa (es. "160 mq")
            - "dettaglio_verniciatura": stringa (es. "Ciclo di vernice intumescente reattiva ad alto spessore per profili aperti pesanti con primer epossidico anticorrosivo")
            - "note_tecniche": stringa (Considerazioni di calcolo rigorose: inclusione delle azioni del vento secondo NTC 2018 con verifica al ribaltamento e suzione di copertura, sezioni maggiorate per margini conservativi.)
            
            Testo da analizzare:
            "{testo_commerciale}"
            """
            
            with st.spinner('L\'ingegnere virtuale sta elaborando i carichi climatici e le verifiche strutturali...'):
                risposta_ia = model.generate_content(prompt)
                testo_risposta = risposta_ia.text.strip()
                if testo_risposta.startswith("```json"):
                    testo_risposta = testo_risposta[7:]
                if testo_risposta.endswith("```"):
                    testo_risposta = testo_risposta[:-3]
                
                dati = json.loads(testo_risposta.strip())
                
                # Salviamo i dati nella sessione di Streamlit per poterli usare per il download
                st.session_state['dati_ultimi'] = dati
                
                st.success("Dimensionamento strutturale e analisi Vento/Neve completati con successo!")
                
        except Exception as e:
            st.error(f"C'è stato un problema nel calcolo strutturale: {e}")

# Se i dati sono stati calcolati, mostriamo la dashboard e il pulsante di download
if 'dati_ultimi' in st.session_state:
    dati = st.session_state['dati_ultimi']
    
    st.markdown("---")
    
    # Pulsante di Download in Word
    col_dl1, col_dl2, col_dl3 = st.columns([1, 2, 1])
    with col_dl2:
        word_file = genera_word_report(dati)
        st.download_button(
            label="📄 Scarica Relazione Tecnica in formato Word (.docx)",
            data=word_file,
            file_name=f"Relazione_Predimensionamento_{dati.get('luogo', 'Progetto').replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
    
    st.markdown("---")
    
    # Sezione 1: Parametri geometrici e climatici (Neve e Vento)
    st.markdown("### 📍 1. Dati geometrici e climatici (NTC 2018)")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Località", dati.get("luogo", "Bolzano"))
    c2.metric("Carico Neve (qsk)", f"{dati.get('qsk', 1.5)} kN/m²")
    c3.metric("Zona Vento", dati.get("zona_vento", "Zona 3"))
    c4.metric("Pressione Vento", dati.get("pressione_vento", "0.80 kN/m²"))
    c5.metric("Interasse Portali", f"{dati.get('interasse_portali', 6.0)} m")
    
    st.markdown("---")
    
    # Sezione 2: Arcarecci
    st.markdown("### 🪵 2. Arcarecci di Copertura")
    st.info(f"**Sezione Consigliata:** {dati.get('sezione_arcarecci', 'N.D.')} | **Stato:** {dati.get('verifica_arcarecci', 'Verificato')}")
    
    st.markdown("---")
    
    # Sezione 3: Travi Principali / Portali (Confronto 3 Materiali)
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
    
    # Sezione 4: Pilastri (Confronto 3 Materiali)
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
    
    # Sezione 5: Controventi e Stabilizzazione
    st.markdown("### 🔗 5. Stabilizzazione e Controventi (Legno vs Acciaio)")
    
    col_cv1, col_cv2 = st.columns(2)
    
    with col_cv1:
        st.markdown("#### 🛡️ Controventi di Copertura (Falda)")
        st.write(f"📍 **Posizionamento:** {dati.get('controventi_copertura_pos', 'N.D.')}")
        st.info(f"🌲 **Opzione Legno:** {dati.get('controventi_copertura_legno', 'N.D.')}")
        st.info(f"⚙️ **Opzione Acciaio:** {dati.get('controventi_copertura_acciaio', 'N.D.')}")
    
    with col_cv2:
        st.markdown("#### 🧱 Controventi di Parete (Baraccatura)")
        st.write(f"📍 **Posizionamento:** {dati.get('controventi_parete_pos', 'N.D.')}")
        dati.get('controventi_parete_legno', 'N.D.')
        st.warning(f"🌲 **Opzione Legno:** {dati.get('controventi_parete_legno', 'N.D.')}")
        st.warning(f"⚙️ **Opzione Acciaio:** {dati.get('controventi_parete_acciaio', 'N.D.')}")
    
    st.markdown("---")
    
    # Sezione 6: Dettaglio Connessioni e Nodi
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
    
    # Sezione 7: Protezione Antincendio e Vernice Intumescente
    st.markdown("### 🔥 7. Requisiti di Resistenza al Fuoco e Vernice Intumescente")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        st.metric("Classe di Resistenza Richiesta", dati.get('classe_resistenza_fuoco', 'R 60'))
    with col_f2:
        st.metric("Superficie Acciaio da Trattare", dati.get('mq_intumescente', 'Non specificato'))
    st.info(f"**Specifiche Ciclo Antincendio:** {dati.get('dettaglio_verniciatura', 'N.D.')}")
    
    st.markdown("---")
    st.markdown("### 📝 8. Relazione e Note Tecniche")
    st.write(dati.get("note_tecniche", "Nessuna nota aggiuntiva."))
