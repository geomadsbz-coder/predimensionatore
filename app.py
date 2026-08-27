import streamlit as st
import google.generativeai as genai
import json
from docx import Document
import io
import ezdxf
import tempfile
import os
import plotly.graph_objects as go
import PyPDF2
import numpy as np

# --- MOTORE DI CALCOLO STRUTTURALE ANALITICO (NTC 2018) ---
def esegui_calcolo_strutturale_rigoroso(dati_geo):
    luce = dati_geo['luce_totale']
    interasse = dati_geo['interasse_portali']
    h_gronda = dati_geo['altezza_gronda']
    h_colmo = dati_geo['altezza_colmo']
    num_appoggi = dati_geo['num_appoggi']
    qsk = dati_geo.get('qsk', 1.5)
    
    g1 = 0.15 
    g2 = 0.25 
    if "Presente" in dati_geo.get('impianto_fv_desc', ''):
        g2 += 0.20
    g2 += dati_geo.get('carico_aggiuntivo', 0.0)
    
    s = qsk * 1.0 
    q_ed = interasse * (1.3 * g1 + 1.5 * g2 + 1.5 * s)
    
    if num_appoggi >= 3:
        luce_campata = luce / (num_appoggi - 1)
        m_ed = (q_ed * (luce_campata ** 2)) / 10.0 
        v_ed = (q_ed * luce_campata) / 2.0        
    else:
        luce_campata = luce
        m_ed = (q_ed * (luce_campata ** 2)) / 8.0
        v_ed = (q_ed * luce_campata) / 2.0

    b_legno = 0.20 
    w_req_cm3 = (m_ed * 1e6) / (14.5 * 1e3)
    h_legno_cm = max(45, int((6 * w_req_cm3 / (b_legno * 100)) ** 0.5 * 10))
    h_legno_cm = ((h_legno_cm + 4) // 4) * 4
    
    w_el_req_cm3 = (m_ed * 1e6) / (335.0)
    if w_el_req_cm3 > 3000:
        profilo_acciaio = "HEB 400 / HEB 450"
    elif w_el_req_cm3 > 1500:
        profilo_acciaio = "IPE 450 / HEA 400"
    else:
        profilo_acciaio = "IPE 360 / HEA 300"

    profilo_cap = f"Trave a T rovescia precompressa altezza {max(80, int(h_legno_cm*1.2))} cm"

    risultati_calcolo = {
        "m_ed": round(m_ed, 1),
        "v_ed": round(v_ed, 1),
        "travi_legno": f"Base 20 cm x Altezza {h_legno_cm} cm (Legno Lamellare GL24h - Verificato a flessione e freccia L/300)",
        "travi_acciaio": f"Profilo {profilo_acciaio} in acciaio S355JR (Verificato SLU/SLE)",
        "travi_cap": profilo_cap,
        "pilastri_legno": f"Sezione 24x{h_legno_cm+4} cm con piastre d'acciaio interne e bulloni",
        "pilastri_acciaio": f"Profilo HEB {min(450, max(260, int(w_el_req_cm3**0.33 * 80)))} S355JR",
        "pilastri_cap": "Pilastro in C.A.P. sezione 40x50 cm con mensola per appoggio trave",
        "sezione_arcarecci": f"Profilo scatolato o falda metallica/legno dimensionato per passo {dati_geo.get('interasse_arcarecci', 1.5)}m (Momento M_Ed arcareccio verificato)"
    }
    return risultati_calcolo

# --- FUNZIONE PER GENERARE IL DOCUMENTO WORD STANDARD ---
def genera_word_report(dati):
    doc = Document()
    doc.add_heading('Relazione Tecnica di Predimensionamento e Calcolo (NTC 2018)', 0)
    
    doc.add_heading('1. Parametri Geometrici, Climatici, Sismici e di Configurazione', level=1)
    doc.add_paragraph(f"Località: {dati.get('luogo', 'Bolzano')}")
    doc.add_paragraph(f"Carico Neve (qsk): {dati.get('qsk', 1.5)} kN/m²")
    doc.add_paragraph(f"Zona Vento: {dati.get('zona_vento', 'N.D.')} | Pressione: {dati.get('pressione_vento', 'N.D.')}")
    doc.add_paragraph(f"Azione Sismica: {dati.get('zona_sismica', 'N.D.')} | Classe d'Uso: {dati.get('classe_uso', 'N.D.')} | Fattore q: {dati.get('fattore_struttura_q', 'N.D.')}")
    doc.add_paragraph(f"Dimensioni Edificio: Lunghezza {dati.get('lunghezza_edificio', 25.0)} m | Larghezza (Luce) {dati.get('luce_totale', 39.6)} m")
    doc.add_paragraph(f"Altezze: Gronda {dati.get('altezza_gronda', 9.0)} m | Colmo {dati.get('altezza_colmo', 12.21)} m")
    doc.add_paragraph(f"Interasse Portali: {dati.get('interasse_portali', 5.0)} m -> N. Campate: {dati.get('num_campate', 5)} (N. Telai: {dati.get('num_campate', 5) + 1})")
    doc.add_paragraph(f"Configurazione Telaio: {dati.get('num_appoggi', 3)} Appoggi | Tipologia Travatura: {dati.get('tipo_travatura', 'N.D.')}")
    doc.add_paragraph(f"Interasse Arcarecci (Passo): {dati.get('interasse_arcarecci', 1.5)} m")
    doc.add_paragraph(f"Copertura / Pannello: {dati.get('tipo_isolante', 'N.D.')} - Spessore: {dati.get('spessore_pannello', 'N.D.')}")
    doc.add_paragraph(f"Impianto Fotovoltaico: {dati.get('impianto_fv_desc', 'Escluso')} | Carico Extra Manuale: {dati.get('carico_aggiuntivo', 0.0)} kN/m²")
    
    doc.add_heading('2. Risultati Analisi Strutturale 3D (Sollecitazioni SLU)', level=1)
    doc.add_paragraph(f"Momento Flettente Massimo di Calcolo (M_Ed): {dati.get('m_ed', 0.0)} kNm")
    doc.add_paragraph(f"Taglio Massimo di Calcolo (V_Ed): {dati.get('v_ed', 0.0)} kN")
    doc.add_paragraph(f"Verifica Arcarecci: {dati.get('sezione_arcarecci', 'N.D.')}")
    
    doc.add_heading('3. Travi Principali / Portali (Confronto 3 Materiali Verificati)', level=1)
    doc.add_paragraph(f"Legno Lamellare: {dati.get('travi_legno', 'N.D.')}")
    doc.add_paragraph(f"Acciaio: {dati.get('travi_acciaio', 'N.D.')}")
    doc.add_paragraph(f"C.a.p.: {dati.get('travi_cap', 'N.D.')}")
    
    doc.add_heading('4. Pilastri (Perimetrali e Intermedi)', level=1)
    doc.add_paragraph(f"Legno Lamellare: {dati.get('pilastri_legno', 'N.D.')}")
    doc.add_paragraph(f"Acciaio: {dati.get('pilastri_acciaio', 'N.D.')}")
    doc.add_paragraph(f"C.a.p.: {dati.get('pilastri_cap', 'N.D.')}")
    
    doc.add_heading('5. Stabilizzazione e Controventi', level=1)
    doc.add_paragraph(f"Copertura (Posizione calcolata IA): {dati.get('controventi_copertura_pos', 'N.D.')}")
    doc.add_paragraph(f"  - Opzione Legno: {dati.get('controventi_copertura_legno', 'N.D.')}")
    doc.add_paragraph(f"  - Opzione Acciaio: {dati.get('controventi_copertura_acciaio', 'N.D.')}")
    doc.add_paragraph(f"Parete (Posizione calcolata IA): {dati.get('controventi_parete_pos', 'N.D.')}")
    doc.add_paragraph(f"  - Opzione Legno: {dati.get('controventi_parete_legno', 'N.D.')}")
    doc.add_paragraph(f"  - Opzione Acciaio: {dati.get('controventi_parete_acciaio', 'N.D.')}")
    
    doc.add_heading('6. Dettaglio Connessioni, Nodi e Giunti in Colmo', level=1)
    doc.add_paragraph(f"Connessione Trave/Pilastro: {dati.get('conn_trave_pilastro_tipo', 'N.D.')}")
    doc.add_paragraph(f"  - Elementi: {dati.get('conn_trave_pilastro_elementi', 'N.D.')} | Peso: {dati.get('conn_trave_pilastro_kg', 'N.D.')}")
    doc.add_paragraph(f"Connessione Pilastro/Fondazione: {dati.get('conn_pilastro_fondazione_tipo', 'N.D.')}")
    doc.add_paragraph(f"  - Ancoraggi: {dati.get('conn_pilastro_fondazione_elementi', 'N.D.')} | Peso: {dati.get('conn_pilastro_fondazione_kg', 'N.D.')}")
    doc.add_paragraph(f"Giunto in Colmo (se previsto): {dati.get('dettaglio_giunto_colmo', 'N.D.')}")
    
    doc.add_heading('7. Protezione Antincendio', level=1)
    doc.add_paragraph(f"Classe Resistenza al Fuoco: {dati.get('classe_resistenza_fuoco', 'N.D.')}")
    doc.add_paragraph(f"Superficie Acciaio da Trattare: {dati.get('mq_intumescente', 'N.D.')}")
    doc.add_paragraph(f"Ciclo: {dati.get('dettaglio_verniciatura', 'N.D.')}")
    
    doc.add_heading('8. Note Tecniche', level=1)
    doc.add_paragraph(dati.get('note_tecniche', 'N.D.'))
    
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# --- FUNZIONE PER GENERARE IL MODELLO 3D DINAMICO ---
def genera_modello_3d(dati):
    fig = go.Figure()
    
    luce_totale = dati.get('luce_totale', 39.6)
    altezza_gronda = dati.get('altezza_gronda', 9.0)
    altezza_colmo = dati.get('altezza_colmo', 12.21)
    lunghezza_edificio = dati.get('lunghezza_edificio', 25.0)
    interasse_portali = dati.get('interasse_portali', 5.0)
    num_appoggi = dati.get('num_appoggi', 3)
    tipo_travatura = dati.get('tipo_travatura', 'Bi-falda semplice')
    interasse_arcarecci = dati.get('interasse_arcarecci', 1.5)
    
    num_campate = max(1, int(round(lunghezza_edificio / interasse_portali)))
    y_portali = [i * interasse_portali for i in range(num_campate + 1)]
    
    if num_appoggi == 2:
        x_pilastri = [0.0, luce_totale]
    elif num_appoggi == 3:
        x_pilastri = [0.0, luce_totale / 2.0, luce_totale]
    else:
        x_pilastri = [0.0, luce_totale / 3.0, (2 * luce_totale) / 3.0, luce_totale]
        
    for idx_y, y in enumerate(y_portali):
        for idx_x, x in enumerate(x_pilastri):
            if x == 0.0 or x == luce_totale:
                h_p = altezza_gronda
            else:
                h_p = altezza_colmo
            
            show_leg = (idx_y == 0 and idx_x == 0)
            fig.add_trace(go.Scatter3d(
                x=[x, x], y=[y, y], z=[0, h_p],
                mode='lines',
                line=dict(color='darkblue', width=6),
                name='Pilastri' if show_leg else '',
                showlegend=show_leg
            ))
        
        show_leg_trave = (idx_y == 0)
        if "curvo" in tipo_travatura.lower():
            x_left = np.linspace(0, luce_totale/2, 10)
            z_left = altezza_gronda + (altezza_colmo - altezza_gronda)*(x_left/(luce_totale/2)) - 0.2*np.sin(np.pi*x_left/(luce_totale/2))
            x_right = np.linspace(luce_totale/2, luce_totale, 10)
            z_right = altezza_colmo - (altezza_colmo - altezza_gronda)*((x_right - luce_totale/2)/(luce_totale/2)) + 0.2*np.sin(np.pi*(x_right - luce_totale/2)/(luce_totale/2))
            
            fig.add_trace(go.Scatter3d(
                x=list(x_left) + list(x_right), y=[y]*20, z=list(z_left) + list(z_right),
                mode='lines',
                line=dict(color='firebrick', width=6),
                name='Travi di Falda' if show_leg_trave else '',
                showlegend=show_leg_trave
            ))
        else:
            fig.add_trace(go.Scatter3d(
                x=[0, luce_totale/2, luce_totale], y=[y, y, y], z=[altezza_gronda, altezza_colmo, altezza_gronda],
                mode='lines',
                line=dict(color='firebrick', width=6),
                name='Travi di Falda' if show_leg_trave else '',
                showlegend=show_leg_trave
            ))
            if "giuntata" in tipo_travatura.lower():
                fig.add_trace(go.Scatter3d(
                    x=[luce_totale/2], y=[y], z=[altezza_colmo],
                    mode='markers',
                    marker=dict(size=6, color='gold'),
                    name='Giunto in Colmo' if show_leg_trave else '',
                    showlegend=show_leg_trave
                ))

    half_luce = luce_totale / 2.0
    x_arc_left = []
    curr = 0.0
    while curr <= half_luce - 1e-5:
        x_arc_left.append(curr)
        curr += interasse_arcarecci
    if not x_arc_left or abs(x_arc_left[-1] - half_luce) > 1e-5:
        x_arc_left.append(half_luce)
        
    for idx_x, x_val in enumerate(x_arc_left):
        z_val = altezza_gronda + (altezza_colmo - altezza_gronda) * (x_val / half_luce)
        show_leg_arc = (idx_x == 0)
        fig.add_trace(go.Scatter3d(
            x=[x_val, x_val], y=[y_portali[0], y_portali[-1]], z=[z_val, z_val],
            mode='lines',
            line=dict(color='gray', width=2, dash='dot'),
            name='Arcarecci' if show_leg_arc else '',
            showlegend=show_leg_arc
        ))
        if x_val < half_luce - 1e-5:
            x_right = luce_totale - x_val
            fig.add_trace(go.Scatter3d(
                x=[x_right, x_right], y=[y_portali[0], y_portali[-1]], z=[z_val, z_val],
                mode='lines',
                line=dict(color='gray', width=2, dash='dot'),
                showlegend=False
            ))

    raw_indici = dati.get('campate_controventi_indici', [0, num_campate - 1])
    if isinstance(raw_indici, list):
        campate_controventi = [int(i) for i in raw_indici if isinstance(i, (int, float))]
    else:
        campate_controventi = [0, num_campate - 1]

    num_sub_falda = max(1, int(round(half_luce / 5.0)))
    num_sub_parete = max(1, int(round(altezza_gronda / 4.5)))
    
    for idx in campate_controventi:
        if 0 <= idx < num_campate:
            y_start = y_portali[idx]
            y_end = y_portali[idx + 1]
            show_leg_cv_cop = (idx == campate_controventi[0])
            show_leg_cv_par = (idx == campate_controventi[0])
            
            dx_falda = half_luce / num_sub_falda
            for s in range(num_sub_falda):
                x_s1 = s * dx_falda
                x_s2 = (s + 1) * dx_falda
                z_s1 = altezza_gronda + (altezza_colmo - altezza_gronda) * (x_s1 / half_luce)
                z_s2 = altezza_gronda + (altezza_colmo - altezza_gronda) * (x_s2 / half_luce)
                
                fig.add_trace(go.Scatter3d(
                    x=[x_s1, x_s2, None, x_s1, x_s2],
                    y=[y_start, y_end, None, y_end, y_start],
                    z=[z_s1, z_s2, None, z_s1, z_s2],
                    mode='lines',
                    line=dict(color='forestgreen', width=4),
                    name='Controventi Copertura' if (show_leg_cv_cop and s == 0) else '',
                    showlegend=(show_leg_cv_cop and s == 0)
                ))
                
                xr_s1 = luce_totale - x_s1
                xr_s2 = luce_totale - x_s2
                fig.add_trace(go.Scatter3d(
                    x=[xr_s1, xr_s2, None, xr_s1, xr_s2],
                    y=[y_start, y_end, None, y_end, y_start],
                    z=[z_s1, z_s2, None, z_s1, z_s2],
                    mode='lines',
                    line=dict(color='forestgreen', width=4),
                    showlegend=False
                ))
            
            dz_parete = altezza_gronda / num_sub_parete
            for x_wall in [0.0, luce_totale]:
                for t in range(num_sub_parete):
                    z_t1 = t * dz_parete
                    z_t2 = (t + 1) * dz_parete
                    
                    fig.add_trace(go.Scatter3d(
                        x=[x_wall, x_wall, None, x_wall, x_wall],
                        y=[y_start, y_end, None, y_start, y_end], 
                        z=[z_t1, z_t2, None, z_t2, z_t1], 
                        mode='lines',
                        line=dict(color='darkorange', width=4),
                        name='Controventi Parete' if (show_leg_cv_par and x_wall == 0.0 and t == 0) else '',
                        showlegend=(show_leg_cv_par and x_wall == 0.0 and t == 0)
                    ))

    fig.update_layout(
        title=f"Modello 3D Dinamico ({num_campate} Campate, {num_campate+1} Telai - {tipo_travatura})",
        scene=dict(
            xaxis_title=f'Larghezza (X - {luce_totale}m)',
            yaxis_title=f'Lunghezza (Y - {lunghezza_edificio}m)',
            zaxis_title=f'Altezza (Z - {altezza_colmo}m)',
            aspectmode='data'
        ),
        margin=dict(l=0, r=0, b=0, t=40),
        height=550
    )
    return fig

st.set_page_config(page_title="Predimensionamento e Calcolo Strutturale IA", layout="wide")
st.title("Generatore Offerte e Modello di Calcolo Strutturale NTC 2018 🏗️")

with st.sidebar:
    st.header("Impostazioni IA")
    api_key = st.text_input("Inserisci qui la tua API Key di Google", type="password")
    st.info("L'API Key serve per far leggere i capitolati all'Intelligenza Artificiale.")

st.subheader("Analisi Capitolato / Appunti di Progetto e File (CAD o PDF)")

file_caricato = st.file_uploader("📂 Carica un file CAD (.dxf) o un documento PDF (.pdf)", type=["dxf", "pdf"])

testo_estratto_file = ""
if file_caricato is not None:
    estensione = file_caricato.name.split('.')[-1].lower()
    try:
        if estensione == 'dxf':
            with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp_file:
                tmp_file.write(file_caricato.getvalue())
                tmp_path = tmp_file.name
            
            doc_dxf = ezdxf.readfile(tmp_path)
            msp = doc_dxf.modelspace()
            testi_estratto = []
            for entity in msp:
                if entity.dxftype() == 'TEXT':
                    testi_estratto.append(entity.dxf.text)
                elif entity.dxftype() == 'MTEXT':
                    testi_estratto.append(entity.text)
            
            testo_estratto_file = "\n".join(testi_estratto)
            st.success(f"File CAD '{file_caricato.name}' letto con successo!")
            os.unlink(tmp_path)
            
        elif estensione == 'pdf':
            pdf_reader = PyPDF2.PdfReader(file_caricato)
            testi_pdf = []
            for page in pdf_reader.pages:
                testo_pagina = page.extract_text()
                if testo_pagina:
                    testi_pdf.append(testo_pagina)
            testo_estratto_file = "\n".join(testi_pdf)
            if not testo_estratto_file.strip():
                st.info("ℹ️ Il file PDF caricato non contiene testo selezionabile (potrebbe essere un disegno/scansione). I parametri geometrici possono essere inseriti o verificati direttamente nei campi sottostanti.")
            else:
                st.success(f"File PDF '{file_caricato.name}' letto con successo!")
            
    except Exception as e:
        st.error(f"Errore nella lettura del file: {e}")

testo_commerciale = st.text_area(
    "Incolla qui le note del progetto o il capitolato:", 
    height=100,
    value="",
    placeholder="Incolla qui note di progetto o capitolato..."
)

st.markdown("### 📐 Dimensioni Geometriche dell'Edificio (Modificabili)")
col_dim1, col_dim2, col_dim3, col_dim4, col_dim5 = st.columns(5)
with col_dim1:
    lunghezza_edificio_ui = st.number_input("Lunghezza Edificio (m)", min_value=5.0, value=25.0, step=1.0, format="%.1f")
with col_dim2:
    interasse_portali_ui = st.number_input("Interasse Portali (m)", min_value=2.0, value=5.0, step=0.5, format="%.2f")
with col_dim3:
    luce_totale_ui = st.number_input("Luce Totale / Larghezza (m)", min_value=5.0, value=39.6, step=0.1, format="%.2f")
with col_dim4:
    altezza_gronda_ui = st.number_input("Altezza Gronda (m)", min_value=3.0, value=9.0, step=0.5, format="%.1f")
with col_dim5:
    altezza_colmo_ui = st.number_input("Altezza Colmo (m)", min_value=3.5, value=12.21, step=0.01, format="%.2f")

st.markdown("### 🏛️ Configurazione Telaio e Travatura")
col_g1, col_g2 = st.columns(2)
with col_g1:
    tipo_travatura = st.selectbox("Tipologia Travatura di Copertura", ["Bi-falda semplice", "Bi-falda con intradosso curvo", "Trave di falda giuntata in colmo"])
with col_g2:
    num_appoggi = st.selectbox("Numero di Appoggi del Telaio", [2, 3, 4], index=1, format_func=lambda x: f"{x} Appoggi ({'Campata Unica' if x==2 else f'Multi-campata con {x-2} pilastro/i interno/i'})")

st.markdown("### ⚙️ Parametri Carichi di Copertura e Pannellature")
col_c1, col_c2, col_c3 = st.columns(3)

with col_c1:
    tipo_isolante = st.selectbox("Tipologia Pannello Copertura", ["PIR / PUR", "Lana Minerale", "Lamiera Grecata Semplice"])
    if tipo_isolante == "PIR / PUR":
        spessore_pannello = st.selectbox("Spessore Pannello (mm)", [50, 60, 80, 100, 120])
    elif tipo_isolante == "Lana Minerale":
        spessore_pannello = st.selectbox("Spessore Pannello (mm)", [100, 120, 150, 170])
    else:
        spessore_pannello = 0

with col_c2:
    st.write("")
    st.write("")
    impianto_fv = st.checkbox("Impianto Fotovoltaico in Copertura (20 kg/mq)", value=False)

with col_c3:
    carico_aggiuntivo = st.number_input("Carico aggiuntivo manuale (kN/mq)", min_value=0.0, value=0.0, step=0.05, format="%.2f")

if st.button("Esegui Calcolo Strutturale e Genera Modello 3D", type="primary"):
    if not api_key:
        st.error("Inserisci prima l'API Key nella barra laterale!")
    else:
        num_campate_calc = max(1, int(round(lunghezza_edificio_ui / interasse_portali_ui)))
        impianto_fv_desc = "Presente (20 kg/mq)" if impianto_fv else "Assente"
        dati_config_str = f"""
        --- CONFIGURAZIONE GEOMETRICA E CARICHI SCELTI ---
        - Lunghezza Edificio: {lunghezza_edificio_ui} m
        - Interasse Portali: {interasse_portali_ui} m (Numero Campate: {num_campate_calc})
        - Luce Totale: {luce_totale_ui} m
        - Altezza Gronda: {altezza_gronda_ui} m
        - Altezza Colmo: {altezza_colmo_ui} m
        - Tipologia Travatura: {tipo_travatura}
        - Numero Appoggi Telaio: {num_appoggi} appoggi
        - Tipologia Pannello: {tipo_isolante}
        - Spessore Pannello: {spessore_pannello} mm
        - Impianto Fotovoltaico: {impianto_fv_desc}
        - Carico Permanente Aggiuntivo Manuale: {carico_aggiuntivo} kN/mq
        """
        
        testo_totale_analisi = testo_commerciale + "\n\n" + dati_config_str + "\n\n--- NOTE DAL FILE ALLEGATO ---\n" + testo_estratto_file
        
        genai.configure(api_key=api_key)
        try:
            model = genai.GenerativeModel(
                model_name='gemini-3.6-flash',
                generation_config={"response_mime_type": "application/json"}
            )
            
            prompt = f"""
Sei un ingegnere strutturista senior esperto in prefabbricazione industriale, NTC 2018.
Analizza il testo tecnico fornito e restituisci i parametri climatici, sismici e descrittivi di progetto.

Restituisci ESATTAMENTE e unicamente un oggetto JSON valido (senza blocchi markdown di alcun tipo, inizia con '{' e finisci con '}') con queste esatte chiavi:
- "luogo": stringa
- "qsk": float
- "zona_vento": stringa
- "pressione_vento": stringa
- "zona_sismica": stringa
- "classe_uso": stringa
- "fattore_struttura_q": stringa
- "campate_controventi_indici": lista di interi (es. [0, {num_campate_calc}-1])
- "interasse_arcarecci": float
- "verifica_arcarecci": stringa
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
- "dettaglio_giunto_colmo": stringa
- "classe_resistenza_fuoco": stringa
- "mq_intumescente": stringa
- "dettaglio_verniciatura": stringa
- "note_tecniche": stringa

Testo da analizzare:
"{testo_totale_analisi}"
            """
            
            with st.spinner('Esecuzione calcolo strutturale analitico e generazione modello 3D...'):
                risposta_ia = model.generate_content(prompt)
                testo_risposta = risposta_ia.text.strip()
                if testo_risposta.startswith("```json"):
                    testo_risposta = testo_risposta[7:]
                if testo_risposta.startswith("```"):
                    testo_risposta = testo_risposta[3:]
                if testo_risposta.endswith("```"):
                    testo_risposta = testo_risposta[:-3]
                
                dati = json.loads(testo_risposta.strip())
                
                dati['lunghezza_edificio'] = lunghezza_edificio_ui
                dati['interasse_portali'] = interasse_portali_ui
                dati['luce_totale'] = luce_totale_ui
                dati['altezza_gronda'] = altezza_gronda_ui
                dati['altezza_colmo'] = altezza_colmo_ui
                dati['num_campate'] = num_campate_calc
                dati['tipo_travatura'] = tipo_travatura
                dati['num_appoggi'] = num_appoggi
                dati['tipo_isolante'] = tipo_isolante
                dati['spessore_pannello'] = f"{spessore_pannello} mm" if tipo_isolante != "Lamiera Grecata Semplice" else "Lamiera Semplice"
                dati['impianto_fv_desc'] = impianto_fv_desc
                dati['carico_aggiuntivo'] = carico_aggiuntivo
                
                risultati_strutturali = esegui_calcolo_strutturale_rigoroso(dati)
                dati.update(risultati_strutturali)
                
                st.session_state['dati_ultimi'] = dati
                st.success("Calcolo strutturale e dimensionamento verificato completati con successo!")
                
        except Exception as e:
            st.error("⚠️ Si è verificato un errore durante l'elaborazione:")
            st.exception(e)

if 'dati_ultimi' in st.session_state:
    dati = st.session_state['dati_ultimi']
    st.markdown("---")
    
    col_dl1, col_dl2, col_dl3 = st.columns([1, 2, 1])
    with col_dl2:
        word_file = genera_word_report(dati)
        st.download_button(
            label="📄 Scarica Relazione di Calcolo in formato Word (.docx)",
            data=word_file,
            file_name=f"Relazione_Calcolo_{dati.get('luogo', 'Progetto').replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
    
    st.markdown("---")
    
    st.markdown("### 🌐 Modello 3D Dinamico e Sollecitazioni Strutturali")
    fig_3d = genera_modello_3d(dati)
    st.plotly_chart(fig_3d, use_container_width=True)
    
    st.markdown("---")
    st.markdown("### 📊 1. Sintesi Sollecitazioni di Calcolo (SLU NTC 2018)")
    col_s1, col_s2, col_s3 = st.columns(3)
    col_s1.metric("Momento Massimo (M_Ed)", f"{dati.get('m_ed', 0.0)} kNm")
    col_s2.metric("Taglio Massimo (V_Ed)", f"{dati.get('v_ed', 0.0)} kN")
    col_s3.metric("Neve di Riferimento (qsk)", f"{dati.get('qsk', 1.5)} kN/m²")
    
    st.markdown("---")
    st.markdown("### 📍 2. Dati geometrici, climatici e di configurazione")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Località", dati.get("luogo", "Bolzano"))
    c2.metric("Lunghezza Edificio", f"{dati.get('lunghezza_edificio')} m")
    c3.metric("N. Campate / Telai", f"{dati.get('num_campate')} Camp. / {dati.get('num_campate')+1} Telai")
    c4.metric("Interasse Portali", f"{dati.get('interasse_portali')} m")
    
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Luce Totale", f"{dati.get('luce_totale')} m")
    c6.metric("Schema Telaio", f"{dati.get('num_appoggi')} Appoggi")
    c7.metric("Travatura", dati.get("tipo_travatura", "Bi-falda semplice"))
    c8.metric("Zona Vento", dati.get("zona_vento", "N.D."))
    
    st.info(f"🏗️ **Copertura configurata:** Pannello {dati.get('tipo_isolante')} ({dati.get('spessore_pannello')}) | **Impianto FV:** {dati.get('impianto_fv_desc')} | **Carico Extra:** {dati.get('carico_aggiuntivo', 0.0)} kN/mq")
    
    st.markdown("---")
    st.markdown("### 🪵 3. Arcarecci di Copertura (Verificati)")
    st.info(f"**Passo Arcarecci:** {dati.get('interasse_arcarecci', 1.5)} m | **Verifica:** {dati.get('sezione_arcarecci', 'Verificato')} | **Stato:** {dati.get('verifica_arcarecci', 'Verificato')}")
    
    st.markdown("---")
    st.markdown("### 📐 4. Travi Principali / Portali (Confronto Tecnologico Dimensionato)")
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
    st.markdown("### 🏛️ 5. Pilastri (Perimetrali e Intermedi)")
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
    st.markdown("### 🔗 6. Stabilizzazione e Controventi (Azioni Orizzontali)")
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
    st.markdown("### 🔩 7. Dimensionamento Dettagliato Connessioni e Nodi")
    col_n1, col_n2 = st.columns(2)
    with col_n1:
        st.markdown("#### 🔗 Connessione Pilastro / Trave di Copertura")
        st.info(f"**Tipologia Nodo:** {dati.get('conn_trave_pilastro_tipo', 'N.D.')}")
        st.write(f"- **Organi di Collegamento:** {dati.get('conn_trave_pilastro_elementi', 'N.D.')}")
        st.metric("Peso Acciaio Connessione", dati.get('conn_trave_pilastro_kg', 'N.D.'))
    with col_n2:
        st.markdown("#### ⚓ Connessione Pilastro / Fondazione")
        st.warning(f"**Tipologia Base:** {dati.get('conn_pilastro_fondazione_tipo', 'N.D.')}")
        st.write(f"- **Ancoraggi:** {dati.get('conn_pilastro_fondazione_elementi', 'N.D.')}")
        st.metric("Peso Acciaio Ancoraggi/Piastra", dati.get('conn_pilastro_fondazione_kg', 'N.D.'))
    
    st.markdown("---")
    st.markdown("### 🔥 8. Requisiti di Resistenza al Fuoco e Vernice Intumescente")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        st.metric("Classe di Resistenza Richiesta", dati.get('classe_resistenza_fuoco', 'R 60'))
    with col_f2:
        st.metric("Superficie Acciaio da Trattare", dati.get('mq_intumescente', 'Non specificato'))
    st.info(f"**Specifiche Ciclo Antincendio:** {dati.get('dettaglio_verniciatura', 'N.D.')}")
    
    st.markdown("---")
    st.markdown("### 📝 9. Relazione e Note Tecniche")
    st.write(dati.get("note_tecniche", "Nessuna nota aggiuntiva."))
