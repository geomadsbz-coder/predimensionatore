[app.py](https://github.com/user-attachments/files/31502905/app.py)
import streamlit as st
import google.generativeai as genai
import json

# --- 1. MOTORE DI CALCOLO STRUTTURALE ---
def calcola_carico_neve(qsk, mu_i=0.8, c_e=1.0, c_t=1.0):
    return round(mu_i * qsk * c_e * c_t, 2)

def calcola_momento_flettente(carico_lineare, luce):
    return round((carico_lineare * (luce ** 2)) / 8, 2)

# --- 2. CONFIGURAZIONE INTERFACCIA ---
st.set_page_config(page_title="Predimensionamento IA", layout="wide")
st.title("Generatore Offerte Tecniche con IA 🏗️")

# Barra laterale per la sicurezza
with st.sidebar:
    st.header("Impostazioni IA")
    api_key = st.text_input("Inserisci qui la tua API Key di Google", type="password")
    st.info("L'API Key serve per far leggere il testo all'Intelligenza Artificiale.")

# --- 3. L'INTERFACCIA DI INPUT TESTUALE ---
st.subheader("Analisi Capitolato / Appunti")
testo_commerciale = st.text_area(
    "Incolla qui le note del progetto:", 
    height=150,
    placeholder="Es. Capannone a Bolzano con interasse portali 5.5m..."
)

# --- 4. LA MAGIA DELL'IA E IL CALCOLO ---
if st.button("Analizza testo e Calcola", type="primary"):
    if not api_key:
        st.error("Inserisci prima l'API Key nella barra laterale!")
    elif non testo_commerciale:
        st.warning("Inserisci del testo da analizzare.")
    else:
        # Configuro l'IA con la tua chiave
        genai.configure(api_key=api_key)
        
        # Scelgo il modello (veloce ed economico per estrazione dati)
        model = genai.GenerativeModel('gemini-1.5-flash', 
                                      generation_config={"response_mime_type": "application/json"})
        
        # Le istruzioni per l'IA (Il Prompt)
        prompt = f"""
        Sei un ingegnere strutturista. Leggi il testo seguente ed estrai i dati necessari al predimensionamento.
        Restituisci ESATTAMENTE un file JSON con queste chiavi:
        "luogo": stringa (nome della città o località)
        "qsk": numero float (carico neve al suolo in kN/m², se non specificato ipotizza 1.50)
        "interasse_portali": numero float (in metri, se non specificato ipotizza 6.0)
        "interasse_arcarecci": numero float (in metri, se non specificato ipotizza 1.5)
        
        Testo da analizzare:
        "{testo_commerciale}"
        """
        
        with st.spinner('L\'IA sta analizzando la richiesta...'):
            try:
                # Invio la richiesta all'IA
                risposta_ia = model.generate_content(prompt)
                # Converto il JSON ricevuto in variabili Python
                dati_estratti = json.loads(risposta_ia.text)
                
                # Eseguo i calcoli strutturali con i dati trovati dall'IA
                carico_neve_mq = calcola_carico_neve(dati_estratti["qsk"])
                carico_lineare = carico_neve_mq * dati_estratti["interasse_arcarecci"]
                momento = calcola_momento_flettente(carico_lineare, dati_estratti["interasse_portali"])
                
                # Mostro i risultati
                st.success("Dati estratti con successo!")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("### Dati Intercettati dall'IA")
                    st.write(f"📍 **Luogo:** {dati_estratti['luogo']}")
                    st.write(f"❄️ **Carico Neve Base (qsk):** {dati_estratti['qsk']} kN/m²")
                    st.write(f"📏 **Interasse Portali:** {dati_estratti['interasse_portali']} m")
                    st.write(f"📏 **Interasse Arcarecci:** {dati_estratti['interasse_arcarecci']} m")
                
                with col2:
                    st.markdown("### Calcolo Strutturale")
                    st.metric("Carico Neve Copertura", f"{carico_neve_mq} kN/m²")
                    st.metric("Carico su arcareccio", f"{carico_lineare:.2f} kN/m")
                    st.metric("Momento Flettente Max", f"{momento} kNm")
                    
            except Exception as e:
                st.error(f"C'è stato un problema nella lettura dei dati: {e}")
