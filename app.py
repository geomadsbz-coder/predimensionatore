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
    st.markdown("**WolfSystem / Tecnico:** Strumenti di calcolo automatico per strutture in legno lamellare e acciaio.")

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
        st.warning("Ingloba del testo da analizzare.")
    else:
        genai.configure(api_key=api_key)
        try:
            model = genai.GenerativeModel(
                model_name='gemini-3.6-flash',
                generation_config={"response_mime_type": "application/json"}
            )
            
            prompt = f"""
            Sei un ingegnere strutturista esperto in edifici prefabbricati in legno lamellare e acciaio. 
            Analizza il testo tecnico fornito e calcola un predimensionamento strutturale di massima.
            
            Restituisci ESATTAMENTE e unicamente un oggetto JSON valido (senza blocchi di codice markdown attorno) con queste chiavi esatte:
            - "luogo": stringa (località del progetto, se non specificata scrivi "Bolzano")
            - "qsk": float (carico neve al suolo in kN/m², se non specificato 1.50)
            - "interasse_portali": float (in metri, se non specificato 6.0)
            - "interasse_arcarecci": float (in metri, se non specificato 1.5)
            - "luce_arcarecci": float (in metri, massima luce degli arcarecci)
            - "sezione_arcarecci": stringa (es. "Lamellare GL24h - 120x240 mm")
            - "verifica_arcarecci": stringa (es. "Verificato a flessione e freccia")
            - "sezione_travi_tetto": stringa (es. "Lamellare GL24h - 160x520 mm")
            - "sezione_pilastri": stringa (es. "Lamellare GL24h - 200x400 mm o Profilo HEB 200")
            - "controventi_copertura": stringa (es. "Croci di sgheriglio in fune d'acciaio / Tavole")
            - "note_tecniche": stringa (breve sintesi delle considerazioni di calcolo e resistenza al fuoco)
            
            Testo da analizzare:
            "{testo_commerciale}"
            """
            
            with st.spinner('L\'ingegnere virtuale sta elaborando il dimensionamento...'):
                risposta_ia = model.generate_content(prompt)
                testo_risposta = risposta_ia.text.strip()
                if testo_risposta.startswith("```json"):
                    testo_risposta = testo_risposta[7:]
                if testo_risposta.endswith("```"):
                    testo_risposta = testo_risposta[:-3]
                
                dati = json.loads(testo_risposta.strip())
                
                st.success("Dimensionamento strutturale completato con successo!")
                
                # Sezione 1: Parametri geometrici e ambientali
                st.markdown("### 📍 1. Dati geometrici e climatici")
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Località", dati.get("luogo", "Bolzano"))
                c2.metric("Carico Neve Base (qsk)", f"{dati.get('qsk', 1.5)} kN/m²")
                c3.metric("Interasse Portali", f"{dati.get('interasse_portali', 6.0)} m")
                c4.metric("Interasse Arcarecci", f"{dati.get('interasse_arcarecci', 1.5)} m")
                
                st.markdown("---")
                
                # Sezione 2: Dimensionamento Elementi Strutturali
                st.markdown("### 📐 2. Proposta Sezioni e Dimensionamento Elementi")
                
                col_a, col_b = st.columns(2)
                
                with col_a:
                    st.markdown("#### Arcarecci di Copertura")
                    st.info(f"**Sezione Consigliata:** {dati.get('sezione_arcarecci', 'N.D.')}")
                    st.write(f"- **Luce statica max:** {dati.get('luce_arcarecci', 6.0)} m")
                    st.write(f"- **Stato Verifiche:** {dati.get('verifica_arcarecci', 'Verificato')} ")
                    
                    st.markdown("#### Travi Principali / Portali")
                    st.success(f"**Sezione Consigliata:** {dati.get('sezione_travi_tetto', 'N.D.')}")
                
                with col_b:
                    st.markdown("#### Pilastri")
                    st.warning(f"**Sezione Consigliata:** {dati.get('sezione_pilastri', 'N.D.')}")
                    
                    st.markdown("#### Stabilizzazione e Controventi")
                    st.write(f"- **Sistema:** {dati.get('controventi_copertura', 'Presenti')} ")
                
                st.markdown("---")
                st.markdown("### 📝 3. Relazione e Note Tecniche")
                st.write(dati.get("note_tecniche", "Nessuna nota aggiuntiva."))
                    
        except Exception as e:
            st.error(f"C'è stato un problema nel calcolo strutturale: {e}")
