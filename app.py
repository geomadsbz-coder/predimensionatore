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
    st.markdown("**WolfSystem / Tecnico:** Strumenti di calcolo automatico avanzato per strutture, cicli antincendio e nodi di connessione ad alte prestazioni (NTC 2018).")

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

if st.button("Esegui Dimensionamento Conservativo", type="primary"):
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
            Sei un ingegnere strutturista senior esperto in edifici prefabbricati industriali, dimensionamento rigoroso secondo NTC 2018 / Eurocodici e progettazione esecutiva di nodi metallici e in legno lamellare. 
            Analizza il testo tecnico fornito e calcola un predimensionamento strutturale **conservativo e robusto** (tenendo conto di eventuali carichi neve e vento gravosi). 
            Attenzione particolare alle connessioni: evita soluzioni sottodimensionate. Pretendi piastre spesse, tirafondi multipli di grande diametro, spinotti multipli ad alta resistenza e pesi in acciaio realistici per strutture pesanti.
            
            Restituisci ESATTAMENTE e unicamente un oggetto JSON valido (senza blocchi di codice markdown attorno) con queste chiavi esatte:
            - "luogo": stringa (località del progetto, se non specificata scrivi "Bolzano")
            - "qsk": float (carico neve al suolo in kN/m², se non specificato 1.50)
            - "interasse_portali": float (in metri, se non specificato 6.0)
            - "interasse_arcarecci": float (in metri, se non specificato 1.5)
            - "sezione_arcarecci": stringa (es. "Lamellare GL24h - 140x280 mm")
            - "verifica_arcarecci": stringa (es. "Verificato a flessione, taglio e freccia instabile")
            - "travi_legno": stringa (es. "Trave a falda curva GL30h - 180x1120 mm")
            - "travi_acciaio": stringa (es. "Profilo IPE 550 / Traliccio pesante in acciaio S355")
            - "travi_cap": stringa (es. "Trave a tegolo alare precompresso in C.a.p.")
            - "pilastri_legno": stringa (es. "Lamellare GL32c - 280x480 mm")
            - "pilastri_acciaio": stringa (es. "Profilo HEB 300 S355")
            - "pilastri_cap": stringa (es. "Pilastro prefabbricato in C.a.p. sezione 50x50 cm")
            - "controventi_copertura_legno": stringa (es. "Diagonali in legno lamellare GL24h - 160x200 mm")
            - "controventi_copertura_acciaio": stringa (es. "Doppie croci di Sant'Andrea in tondo d'acciaio Ø24 mm con tenditori registrabili")
            - "controventi_copertura_pos": stringa (es. "N° 2 campi di falda completi in corrispondenza delle campate di testata")
            - "controventi_parete_legno": stringa (es. "Baraccatura pesante e diagonali in legno lamellare GL24h - 160x220 mm")
            - "controventi_parete_acciaio": stringa (es. "Diagonali verticali rigide in profilati cavi strutturali RHS 140x140x8 mm")
            - "controventi_parete_pos": stringa (es. "Campate perimetrali di testata e pareti longitudinali intermedie")
            - "conn_trave_pilastro_tipo": stringa (es. "Connessione a momento rigida con doppia lama d'acciaio interna a scomparsa ad alto spessore e piastra di colmo saldata")
            - "conn_trave_pilastro_elementi": stringa (es. "N° 20 spinotti in acciaio ad alta resistenza Ø16 mm disposti su maglia fitta + copripiastre esterne bullonate")
            - "conn_trave_pilastro_kg": stringa (es. "45.0 kg per nodo (configurazione pesante)")
            - "conn_pilastro_fondazione_tipo": stringa (es. "Piastra di base d'acciaio S355 spessore 35 mm irrigidita da n° 4 fazzoletti triangolari per lato e collare di base")
            - "conn_pilastro_fondazione_elementi": stringa (es. "N° 8 tirafondi di ancoraggio ad alta resistenza M30 L=1200 mm in acciaio 8.8 con piastra d'ancoraggio inferiore")
            - "conn_pilastro_fondazione_kg": stringa (es. "95.0 kg per plinto di fondazione")
            - "classe_resistenza_fuoco": stringa (es. "R 60")
            - "mq_intumescente": stringa (es. "160 mq")
            - "dettaglio_verniciatura": stringa (es. "Ciclo di vernice intumescente reattiva ad alto spessore per profili aperti pesanti con primer epossidico anticorrosivo")
            - "note_tecniche": stringa (Considerazioni di calcolo rigorose: sezioni maggiorate per garantire margini di sicurezza conservativi contro carichi neve e vento estremi, nodi irrigiditi e verificati a rottura duttile.)
            
            Testo da analizzare:
            "{testo_commerciale}"
            """
            
            with st.spinner('L\'ingegnere virtuale sta ricalcolando il dimensionamento con criteri di sicurezza conservativi...'):
                risposta_ia = model.generate_content(prompt)
                testo_risposta = risposta_ia.text.strip()
                if testo_risposta.startswith("```json"):
                    testo_risposta = testo_risposta[7:]
                if testo_risposta.endswith("```"):
                    testo_risposta = testo_risposta[:-3]
                
                dati = json.loads(testo_risposta.strip())
                
                st.success("Dimensionamento strutturale conservativo completato con successo!")
                
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
                
                # Sezione 6: Dettaglio Connessioni e Nodi (Rafforzato)
                st.markdown("### 🔩 6. Dimensionamento Dettagliato Connessioni e Nodi (Configurazione Pesante)")
                
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
                    
        except Exception as e:
            st.error(f"C'è stato un problema nel calcolo strutturale: {e}")
