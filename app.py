import streamlit as st
import google.generativeai as genai
import json

# --- CONFIGURAZIONE INTERFACCIA ---
st.set_page_config(page_title="Predimensionamento Strutturale IA", layout="wide")
st.title("Generatore Offerte Tecniche e Dimensionamento IA 🏗️")

with st.sidebar:
    st.header("Impostazioni IA")
    api_key = st.text_input("Inserisci qui la tua API Key di Google", type="password")
    st.info("L'API Key serve per far leggere il capitolato all'Intelligenza Artificiale.")
    st.markdown("---")
    st.markdown("**WolfSystem / Tecnico:** Strumenti di calcolo automatico per strutture, cicli antincendio e nodi di connessione.")

st.subheader("Analisi Capitolato / Appunti di Progetto")
testo_commerciale = st.text_area(
    "Incolla qui le note del progetto o il capitolato:", 
    height=150,
    value="""N° 19 file di arcarecci in legno lamellare qualità industria non impregnato posti ad interasse secondo adeguato calcolo statico, con luce statica massima di 6,00 m, collegati meccanicamente alla struttura sottostante.
N° 4 travi diagonali di testata in legno lamellare qualità industria non impregnato con luce statica massima di 9,00 m, collegate meccanicamente alla struttura sottostante.
Stabilizzazione sul piano orizzontale e verticale della struttura.
Baraccatura di parete composta da elementi in legno posti ad interasse variabile con luce statica variabile.
Tutti gli elementi strutturali in acciaio saranno protetti in modo adeguato per la verifica al fuoco.
Interasse telaio 6,00ml"""
)

if st.button("Esegui Dimensionamento Completo", type="primary"):
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
            Sei un ingegnere strutturista senior esperto in edifici prefabbricati, dettagli costruttivi di nodi strutturali (legno, acciaio, C.a.p.) e protezione antincendio. 
            Analizza il testo tecnico fornito e calcola un predimensionamento strutturale di massima, includendo in modo dettagliato il dimensionamento delle connessioni principali (Pilastro/Trave di copertura e Pilastro/Fondazione), specificando la tipologia costruttiva, gli elementi di fissaggio (spinotti, bulloni, tirafondi) e stimando il relativo peso in kg dell'acciaio impiegato nei nodi.
            
            Restituisci ESATTAMENTE e unicamente un oggetto JSON valido (senza blocchi di codice markdown attorno) con queste chiavi esatte:
            - "luogo": stringa (località del progetto, se non specificata scrivi "Bolzano")
            - "qsk": float (carico neve al suolo in kN/m², se non specificato 1.50)
            - "interasse_portali": float (in metri, se non specificato 6.0)
            - "interasse_arcarecci": float (in metri, se non specificato 1.5)
            - "sezione_arcarecci": stringa (es. "Lamellare GL24h - 120x240 mm")
            - "verifica_arcarecci": stringa (es. "Verificato a flessione e freccia")
            - "travi_legno": stringa (es. "Trave a falda curva GL28h - 160x1000 mm")
            - "travi_acciaio": stringa (es. "Profilo IPE 500 o Traliccio in acciaio S355")
            - "travi_cap": stringa (es. "Trave a T o a tegolo in C.a.p. precompresso")
            - "pilastri_legno": stringa (es. "Lamellare GL30h - 240x400 mm")
            - "pilastri_acciaio": stringa (es. "Profilo HEB 260 S355")
            - "pilastri_cap": stringa (es. "Pilastro prefabbricato in C.a.p. sezione 40x40 cm")
            - "controventi_copertura_legno": stringa (es. "Diagonali in legno lamellare GL24h - 140x160 mm")
            - "controventi_copertura_acciaio": stringa (es. "Croci in tondo d'acciaio Ø20 mm con tenditori")
            - "controventi_copertura_pos": stringa (es. "N° 2 campi di falda in corrispondenza delle campate di testata")
            - "controventi_parete_legno": stringa (es. "Baraccatura e diagonali in legno lamellare GL24h - 140x180 mm")
            - "controventi_parete_acciaio": stringa (es. "Diagonali verticali in profilati cavi d'acciaio RHS 100x100x5 mm")
            - "controventi_parete_pos": stringa (es. "Campate perimetrali di testata e pareti longitudinali")
            - "conn_trave_pilastro_tipo": stringa (es. "Connessione con lama d'acciaio a scomparsa inserita nel colmo/trave e perni passanti")
            - "conn_trave_pilastro_elementi": stringa (es. "N° 12 spinotti in acciaio ad alta resistenza Ø12 mm e piastra di testa saldata")
            - "conn_trave_pilastro_kg": stringa (es. "18.5 kg per nodo")
            - "conn_pilastro_fondazione_tipo": stringa (es. "Piastra di base in acciaio S355 spessore 25 mm con fazzoletti di irrigidimento e bicchiere/tirafondi")
            - "conn_pilastro_fondazione_elementi": stringa (es. "N° 4 tirafondi di ancoraggio M24 L=800 mm in acciaio 8.8 con dima di posa")
            - "conn_pilastro_fondazione_kg": stringa (es. "34.0 kg per plinto")
            - "classe_resistenza_fuoco": stringa (es. "R 60")
            - "mq_intumescente": stringa (es. "135 mq")
            - "dettaglio_verniciatura": stringa (es. "Ciclo intumescente reattivo per profili aperti/chiusi con primer e finitura")
            - "note_tecniche": stringa (breve sintesi delle considerazioni di calcolo, carichi neve/vento e verifiche)
            
            Testo da analizzare:
            "{testo_commerciale}"
            """
            
            with st.spinner('L\'ingegnere virtuale sta calcolando sezioni, protezione al fuoco e dettagli esecutivi delle connessioni...'):
                risposta_ia = model.generate_content(prompt)
                testo_risposta = risposta_ia.text.strip()
                if testo_risposta.startswith("```json"):
                    testo_risposta = testo_risposta[7:]
                if testo_risposta.endswith("
