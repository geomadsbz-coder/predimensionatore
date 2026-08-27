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
    st.markdown("**WolfSystem / Tecnico:** Strumenti di calcolo automatico per strutture in legno lamellare, acciaio, C.a.p. e cicli di protezione al fuoco.")

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
            Sei un ingegnere strutturista esperto in edifici prefabbricati, protezione antincendio e dimensionamento in legno lamellare, acciaio e C.a.p. 
            Analizza il testo tecnico fornito e calcola un predimensionamento strutturale di massima. 
            Inoltre, individua o deduci la classe di resistenza al fuoco richiesta (R 30, R 60 oppure R 90) per gli elementi in acciaio, e calcola in via preliminare la superficie totale in metri quadri (mq) dei profili in acciaio che necessitano di vernice intumescente (moltiplicando lo sviluppo del perimetro dei profili per la lunghezza complessiva stimata degli elementi metallici).
            
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
            - "classe_resistenza_fuoco": stringa (es. "R 60")
            - "mq_intumescente": stringa (es. "135 mq")
            - "dettaglio_verniciatura": stringa (es. "Ciclo di vernice intumescente reattiva spessore calcolato per profili aperti/chiusi con primer anticorrosivo e finitura protettiva")
            - "note_tecniche": stringa (breve sintesi delle considerazioni di calcolo, carichi neve/vento e verifiche)
            
            Testo da analizzare:
            "{testo_commerciale}"
            """
            
            with st.spinner('L\'ingegnere virtuale sta elaborando il dimensionamento e calcolando i mq di vernice intumescente...'):
                risposta_ia = model.generate_content(prompt)
                testo_risposta = risposta_ia.text.strip()
                if testo_risposta.startswith("```json"):
                    testo_risposta = testo_risposta[7:]
                if testo_risposta.endswith("```"):
                    testo_risposta = testo_risposta[:-3]
                
                dati = json.loads(testo_risposta.strip())
                
                st.success("Dimensionamento strutturale e calcolo antincendio completati con successo!")
                
                # Sezione 1: Parametri geometrici e ambientali
                st.markdown("### 📍 1. Dati geometrici e climatici")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Località", dati.get("luogo", "Bolzano"))
                c2.metric("Carico Neve Base (qsk)", f"{dati.get('qsk', 1.5)} kN/m²")
                c3.metric("Interasse Portali", f"{dati.get('interasse_portali', 6.0)} m")
                c4.metric("Interasse Arcarecci", f"{dati.get('interasse_arcarecci', 1.5)} m")
                
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
                    st.warning(f"🌲 **Opzione Legno:** {dati.get('controventi_parete_legno', 'N.D.')}")
                    st.warning(f"⚙️ **Opzione Acciaio:** {dati.get('controventi_parete_acciaio', 'N.D.')}")
                
                st.markdown("---")
                
                # Sezione 6: Protezione Antincendio e Vernice Intumescente
                st.markdown("### 🔥 6. Requisiti di Resistenza al Fuoco e Vernice Intumescente")
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    st.metric("Classe di Resistenza Richiesta", dati.get('classe_resistenza_fuoco', 'R 60'))
                with col_f2:
                    st.metric("Superficie Acciaio da Trattare", dati.get('mq_intumescente', 'Non specificato'))
                st.info(f"**Specifiche Ciclo Antincendio:** {dati.get('dettaglio_verniciatura', 'N.D.')}")
                
                st.markdown("---")
                st.markdown("### 📝 7. Relazione e Note Tecniche")
                st.write(dati.get("note_tecniche", "Nessuna nota aggiuntiva."))
                    
        except Exception as e:
            st.error(f"C'è stato un problema nel calcolo strutturale: {e}")
