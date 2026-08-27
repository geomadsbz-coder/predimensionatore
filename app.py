import streamlit as st
import google.generativeai as genai
import json

def calcola_carico_neve(qsk, mu_i=0.8, c_e=1.0, c_t=1.0):
    return round(mu_i * qsk * c_e * c_t, 2)

def calcola_momento_flettente(carico_lineare, luce):
    return round((carico_lineare * (luce ** 2)) / 8, 2)

st.set_page_config(page_title="Predimensionamento IA", layout="wide")
st.title("Generatore Offerte Tecniche con IA 🏗️")

with st.sidebar:
    st.header("Impostazioni IA")
    api_key = st.text_input("Inserisci qui la tua API Key di Google", type="password")
    st.info("L'API Key serve per far leggere il testo all'Intelligenza Artificiale.")

st.subheader("Analisi Capitolato / Appunti")
testo_commerciale = st.text_area(
    "Incolla qui le note del progetto:", 
    height=150,
    value="""N° 19 file di arcarecci in legno lamellare qualità industria non impregnato posti ad interasse secondo adeguato calcolo statico, con luce statica massima di 6,00 m, collegati meccanicamente alla struttura sottostante.
N° 4 travi diagonali di testata in legno lamellare qualità industria non impregnato con luce statica massima di 9,00 m, collegate meccanicamente alla struttura sottostante.
Stabilizzazione sul piano orizzontale e verticale della struttura.
Baraccatura di parete composta da elementi in legno posti ad interasse variabile con luce statica variabile.
Tutti gli elementi strutturali in acciaio saranno protetti in modo adeguato per la verifica al fuoco.
Interasse telaio 6,00ml"""
)

if st.button("Analizza testo e Calcola", type="primary"):
    if not api_key:
        st.error("Inserisci prima l'API Key nella barra laterale!")
    elif not testo_commerciale:
        st.warning("Inserisci del testo da analizzare.")
    else:
        genai.configure(api_key=api_key)
        try:
            # Modello impostato rigidamente sull'ultima versione richiesta dall'API
            model = genai.GenerativeModel(
                model_name='gemini-3.6-flash',
                generation_config={"response_mime_type": "application/json"}
            )
            
            prompt = f"""
            Sei un ingegnere strutturista. Leggi il testo seguente ed estrai i dati necessari al predimensionamento.
            Restituisci ESATTAMENTE e unicamente un oggetto JSON valido (senza blocchi di codice markdown attorno) con queste chiavi esatte:
            "luogo": stringa (nome della città o località, se non menzionata scrivi "Bolzano")
            "qsk": numero float (carico neve al suolo in kN/m², se non specificato ipotizza 1.50)
            "interasse_portali": numero float (in metri, se non specificato ipotizza 6.0)
            "interasse_arcarecci": float (in metri, se non specificato ipotizza 1.5)
            
            Testo da analizzare:
            "{testo_commerciale}"
            """
            
            with st.spinner('L\'IA sta analizzando la richiesta...'):
                risposta_ia = model.generate_content(prompt)
                testo_risposta = risposta_ia.text.strip()
                if testo_risposta.startswith("```json"):
                    testo_risposta = testo_risposta[7:]
                if testo_risposta.endswith("```"):
                    testo_risposta = testo_risposta[:-3]
                
                dati_estratti = json.loads(testo_risposta.strip())
                
                carico_neve_mq = calcola_carico_neve(dati_estratti["qsk"])
                carico_lineare = carico_neve_mq * dati_estratti["interasse_arcarecci"]
                momento = calcola_momento_flettente(carico_lineare, dati_estratti["interasse_portali"])
                
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
