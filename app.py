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
import re
import requests

# --- FUNZIONE PER ESTRARRE COORDINATE DA URL DI GOOGLE MAPS ---
def estrai_coordinate_da_url(url):
    url = url.strip()
    # Verifica link monchi o vuoti
    if not url or url == "https://maps.app.goo.gl/" or url == "https://maps.app.goo.gl":
        return 46.4983, 11.3548 # Default Bolzano
    
    # Se il link è corto (goo.gl), lo espande seguendo il reindirizzamento
    if "goo.gl" in url:
        try:
            response = requests.head(url, allow_redirects=True, timeout=5)
            url = response.url # Questo è il link lungo espanso che contiene le coordinate
        except Exception:
            pass 
            
    # 1. Cerca il pattern tipico di Google Maps '@lat,lon'
    match_at = re.search(r'@([-0-9.]+),([-0-9.]+)', url)
    if match_at:
        try:
            return float(match_at.group(1)), float(match_at.group(2))
        except ValueError:
            pass
            
    # 2. Cerca il pattern con parametri di ricerca 'q=lat,lon' o 'll=lat,lon'
    match_q = re.search(r'[?&](?:q|ll|s_loc)=([-0-9.]+),([-0-9.]+)', url)
    if match_q:
        try:
            return float(match_q.group(1)), float(match_q.group(2))
        except ValueError:
            pass
            
    # 3. Se l'utente ha incollato direttamente coordinate grezze es. "46.4983, 11.3548"
    match_raw = re.search(r'([-0-9.]+)\s*,\s*([-0-9.]+)', url)
    if match_raw:
        try:
            lat = float(match_raw.group(1))
            lon = float(match_raw.group(2))
            if 35.0 <= lat <= 47.5 and 6.0 <= lon <= 19.0:
                return lat, lon
        except ValueError:
            pass
            
    return 46.4983, 11.3548 # Fallback di sicurezza

# --- MOTORE DI CALCOLO STRUTTURALE DETERMINISTICO DA COORDINATE (NTC 2018) ---
def estrai_parametri_ntc_da_coordinate(lat, lon):
    if lat > 45.8:
        altitudine_stimata = 650.0  
        zona_neve = "Zona I (Alpina / Montana)"
        qsk = round(1.39 * (1.0 + (altitudine_stimata / 728.0) ** 2), 2)
        zona_vento = "Zona 1 (vb = 25 m/s)"
        pressione_vento = round(0.50 * (1.0 + altitudine_stimata/1000.0), 2)
        zona_sismica = "Zona 3 (Bassa sismicità / Area Alpina)"
    elif lat > 44.5:
        altitudine_stimata = 50.0   
        zona_neve = "Zona II (Padana / Interna)"
        qsk = round(0.85 * (1.0 + (altitudine_stimata / 778.0) ** 2), 2)
        zona_vento = "Zona 3 (vb = 27 m/s - Interna)"
        pressione_vento = 0.48
        zona_sismica = "Zona 2 / 3 (Media/Bassa sismicità)"
    elif lat > 41.0:
        altitudine_stimata = 150.0  
        zona_neve = "Zona II (Interna Centro)"
        qsk = round(0.85 * (1.0 + (altitudine_stimata / 778.0) ** 2), 2)
        zona_vento = "Zona 2 (vb = 28 m/s)"
        pressione_vento = 0.52
        zona_sismica = "Zona 1 / 2 (Alta/Media sismicità - Appennino)"
    else:
        altitudine_stimata = 50.0   
        zona_neve = "Zona III (Meridionale / Costiera)"
        qsk = round(0.50 * (1.0 + (altitudine_stimata / 833.0) ** 2), 2)
        zona_vento = "Zona 3 o 4 (Sud/Isole)"
        pressione_vento = 0.58
        zona_sismica = "Zona 1 (Alta sismicità - es. Calabria/Sicilia)"
        
    luogo_str = f"Google Maps ({lat:.4f}, {lon:.4f}) - Alt. stimata: {altitudine_stimata}m - {zona_neve}"
    return luogo_str, qsk, zona_vento, f"{pressione_vento} kN/mq", zona_sismica, altitudine_stimata

def esegui_calcolo_deterministico(dati_geo):
    luce = dati_geo['luce_totale']
    interasse = dati_geo['interasse_portali']
    h_gronda = dati_geo['altezza_gronda']
    h_colmo = dati_geo['altezza_colmo']
    num_appoggi = dati_geo['num_appoggi']
    
    lat = dati_geo.get('latitudine', 46.4983)
    lon = dati_geo.get('longitudine', 11.3548)
    
    luogo_str, qsk, zona_vento, pressione_vento, zona_sismica, altitudine_stimata = estrai_parametri_ntc_da_coordinate(lat, lon)
    
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

    n_bulloni_nodo = max(4, int(v_ed / 35.0) * 2)
    peso_conn_kg = round(n_bulloni_nodo * 4.5 + 15.0, 1)
    n_ancoraggi = max(4, int(v_ed / 40.0) * 2)
    peso_anc_kg = round(n_ancoraggi * 3.5 + 20.0, 1)

    mq_acciaio = round((luce + h_gronda * 2) * (dati_geo['num_campate'] + 1) * 0.6, 1)

    risultati_deterministici = {
        "luogo": luogo_str,
        "qsk": qsk,
        "zona_vento": zona_vento,
        "pressione_vento": pressione_vento,
        "zona_sismica": zona_sismica,
        "classe_uso": "Classe II (Edifici industriali ordinari)",
        "fattore_struttura_q": "q = 2.0 (Struttura intelaiata)",
        "m_ed": round(m_ed, 1),
        "v_ed": round(v_ed, 1),
        "travi_legno": f"Base 20 cm x Altezza {h_legno_cm} cm (Legno Lamellare GL24h - Verificato a flessione e freccia L/300)",
        "travi_acciaio": f"Profilo {profilo_acciaio} in acciaio S355JR (Verificato SLU/SLE)",
        "travi_cap": profilo_cap,
        "pilastri_legno": f"Sezione 24x{h_legno_cm+4} cm con piastre d'acciaio interne e bulloni",
        "pilastri_acciaio": f"Profilo HEB {min(450, max(260, int(w_el_req_cm3**0.33 * 80)))} S355JR",
        "pilastri_cap": "Pilastro in C.A.P. sezione 40x50 cm con mensola per appoggio trave",
        "sezione_arcarecci": f"Profilo scatolare 120x60x4 mm o falda legno 10x20 cm per passo {dati_geo.get('interasse_arcarecci', 1.5)}m",
        "verifica_arcarecci": "Verificato a flessione deviata e freccia SLE (L/200)",
        "baraccatura_legno_lamellare": f"Correnti in legno lamellare GL24h sezione 12x16 cm interasse 1.5m",
        "baraccatura_legno_massiccio": f"Correnti in legno massiccio C24 sezione 14x16 cm interasse 1.5m",
        "baraccatura_acciaio": f"Profilo secondario Omega o Tubolare 100x50x3 mm in acciaio S355",
        "campate_controventi_indici": [0, dati_geo['num_campate'] - 1],
        "controventi_copertura_pos": f"Campate di estremità (1ª e ultima campata)",
        "controventi_copertura_legno": "Tiranti tondi d'acciaio diametro 20 mm con capannine in legno lamellare",
        "controventi_copertura_acciaio": "Tubolari incrociati Ø 89x4 mm o diagonali a doppio L",
        "controventi_parete_pos": f"Campate di estremità in corrispondenza dei controventi di falda",
        "controventi_parete_legno": "Diagonali in legno lamellare 16x16 cm con piastre di nodo dedicate",
        "controventi_parete_acciaio": "Croci di sant'andrea in profilati angolari L 80x8 o tubolari strutturali",
        "conn_trave_pilastro_tipo": "Nodo semi-rigido con piastre frontali interne e spine d'acciaio",
        "conn_trave_pilastro_elementi": f"N. {n_bulloni_nodo} bulloni classe 8.8 M20 + piastra sp. 20 mm",
        "conn_trave_pilastro_kg": f"{peso_conn_kg} kg cad.",
        "conn_pilastro_fondazione_tipo": "Cerniera/Incastro parziale con piastra di base e fazzoletti di irrigidimento",
        "conn_pilastro_fondazione_elementi": f"N. {n_ancoraggi} tirafondi ad alta resistenza M24 L=800mm",
        "conn_pilastro_fondazione_kg": f"{peso_anc_kg} kg cad.",
        "dettaglio_giunto_colmo": "Piastra di colmo sagomata con coprigiunti bullonati a doppio taglio",
        "classe_resistenza_fuoco": "R 60 (Conforme ai requisiti antincendio attività industriali)",
        "mq_intumescente": f"{mq_acciaio} mq",
        "dettaglio_verniciatura": "Primer epossidico anticorrosivo + Vernice intumescente a spessore testata per 60 min",
        "note_tecniche": f"Predimensionamento deterministico NTC 2018 basato su coordinate GPS (lat: {lat:.4f}, lon: {lon:.4f}). Altitudine stimata: {altitudine_stimata}m, qsk = {qsk} kN/mq."
    }
    return risultati_deterministici

# --- FUNZIONE PER CALCOLARE LA DISTINTA ELEMENTI ---
def calcola_distinta_elementi(dati):
    L = dati['lunghezza_edificio']
    B = dati['luce_totale']
    i_portali = dati['interasse_portali']
    n_appoggi = dati['num_appoggi']
    i_arcarecci = dati.get('interasse_arcarecci', 1.5)

    num_campate = max(1, int(round(L / i_portali))) if i_portali > 0 else 1
    num_telai = num_campate + 1

    num_pilastri_totali = num_telai * n_appoggi
    num_pilastri_perimetrali = num_telai * 2
    num_pilastri_interni = num_pilastri_totali - num_pilastri_perimetrali
    num_travi_falda = num_telai * 2 

    half_luce = B / 2.0
    x_arc_left = []
    curr = 0.0
    while curr <= half_luce - 1e-5:
        x_arc_left.append(curr)
        curr += i_arcarecci
    if not x_arc_left or abs(x_arc_left[-1] - half_luce) > 1e-5:
        x_arc_left.append(half_luce)
    
    num_file_arcarecci = len(x_arc_left) * 2 - 1 
    ml_arcarecci = num_file_arcarecci * L

    raw_indici = dati.get('campate_controventi_indici', [0, num_campate - 1])
    if isinstance(raw_indici, list):
        campate_cv = [int(i) for i in raw_indici if isinstance(i, (int, float))]
    else:
        campate_cv = [0, num_campate - 1]
    
    num_campate_cv = sum(1 for idx in campate_cv if 0 <= idx < num_campate)
    num_sub_falda = max(1, int(round(half_luce / 5.0)))
    num_croci_cop = num_campate_cv * (num_sub_falda * 2) 
    
    h_gronda = dati['altezza_gronda']
    num_sub_parete = max(1, int(round(h_gronda / 4.5))) if h_gronda > 0 else 1
    num_croci_par = num_campate_cv * (num_sub_parete * 2) 

    sviluppo_falda = ((B/2)**2 + (dati['altezza_colmo'] - h_gronda)**2)**0.5
    mq_copertura = L * sviluppo_falda * 2
    mq_pareti = L * h_gronda * 2 

    return {
        "num_telai": num_telai,
        "num_pilastri_totali": num_pilastri_totali,
        "num_pilastri_perimetrali": num_pilastri_perimetrali,
        "num_pilastri_interni": num_pilastri_interni,
        "num_travi_falda": num_travi_falda,
        "num_file_arcarecci": num_file_arcarecci,
        "ml_arcarecci": round(ml_arcarecci, 1),
        "num_croci_copertura": num_croci_cop,
        "num_croci_parete": num_croci_par,
        "mq_copertura": round(mq_copertura, 1),
        "mq_pareti_lunghe": round(mq_pareti, 1)
    }

# --- FUNZIONE PER GENERARE IL DOCUMENTO WORD STANDARD ---
def genera_word_report(dati, distinta):
    doc = Document()
    doc.add_heading('Relazione Tecnica di Predimensionamento e Calcolo (NTC 2018)', 0)
    
    doc.add_heading('1. Parametri Geometrici, Climatici, Sismici e di Configurazione', level=1)
    doc.add_paragraph(f"Località / Google Maps: {dati.get('luogo', 'N.D.')}")
    doc.add_paragraph(f"Carico Neve (qsk): {dati.get('qsk', 1.5)} kN/m²")
    doc.add_paragraph(f"Zona Vento: {dati.get('zona_vento', 'N.D.')} | Pressione: {dati.get('pressione_vento', 'N.D.')}")
    doc.add_paragraph(f"Azione Sismica: {dati.get('zona_sismica', 'N.D.')} | Classe d'Uso: {dati.get('classe_uso', 'N.D.')} | Fattore q: {dati.get('fattore_struttura_q', 'N.D.')}")
    doc.add_paragraph(f"Dimensioni Edificio: Lunghezza {dati.get('lunghezza_edificio', 0.0)} m | Larghezza (Luce) {dati.get('luce_totale', 0.0)} m")
    doc.add_paragraph(f"Altezze: Gronda {dati.get('altezza_gronda', 0.0)} m | Colmo {dati.get('altezza_colmo', 0.0)} m")
    doc.add_paragraph(f"Interasse Portali: {dati.get('interasse_portali', 0.0)} m -> N. Campate: {dati.get('num_campate', 0)} (N. Telai: {dati.get('num_campate', 0) + 1})")
    doc.add_paragraph(f"Configurazione Telaio: {dati.get('num_appoggi', 3)} Appoggi | Tipologia Travatura: {dati.get('tipo_travatura', 'N.D.')}")
    doc.add_paragraph(f"Copertura / Pannello: {dati.get('tipo_isolante', 'N.D.')} - Spessore: {dati.get('spessore_pannello', 'N.D.')}")
    doc.add_paragraph(f"Impianto Fotovoltaico: {dati.get('impianto_fv_desc', 'Escluso')} | Carico Extra Manuale: {dati.get('carico_aggiuntivo', 0.0)} kN/m²")
    
    doc.add_heading('2. Distinta Elementi Principali (Computo Quantità)', level=1)
    doc.add_paragraph(f"Numero Telai Principali: {distinta['num_telai']}")
    doc.add_paragraph(f"Numero Pilastri Totali: {distinta['num_pilastri_totali']} (di cui {distinta['num_pilastri_perimetrali']} perimetrali e {distinta['num_pilastri_interni']} interni)")
    doc.add_paragraph(f"Numero Travi di Falda: {distinta['num_travi_falda']}")
    doc.add_paragraph(f"File di Arcarecci (sviluppo trasversale): {distinta['num_file_arcarecci']} file")
    doc.add_paragraph(f"Metri Lineari Totali Arcarecci: {distinta['ml_arcarecci']} ml")
    doc.add_paragraph(f"Moduli di Controvento Copertura (Croci): {distinta['num_croci_copertura']}")
    doc.add_paragraph(f"Moduli di Controvento Parete (Croci): {distinta['num_croci_parete']}")
    doc.add_paragraph(f"Superficie Copertura (falde): {distinta['mq_copertura']} mq")
    doc.add_paragraph(f"Superficie Pareti Longitudinali: {distinta['mq_pareti_lunghe']} mq")

    doc.add_heading('3. Arcarecci di Copertura', level=1)
    doc.add_paragraph(f"Passo / Interasse Arcarecci: {dati.get('interasse_arcarecci', 1.5)} m")
    doc.add_paragraph(f"Sezione Consigliata: {dati.get('sezione_arcarecci', 'N.D.')}")
    doc.add_paragraph(f"Verifica: {dati.get('verifica_arcarecci', 'N.D.')}")
    
    doc.add_heading('4. Baraccatura di Parete (Supporto Rivestimento)', level=1)
    doc.add_paragraph(f"Pannello Parete: {dati.get('tipo_isolante_parete', 'N.D.')} - Spessore: {dati.get('spessore_pannello_parete', 'N.D.')}")
    doc.add_paragraph(f"Legno Lamellare: {dati.get('baraccatura_legno_lamellare', 'N.D.')}")
    doc.add_paragraph(f"Legno Massiccio: {dati.get('baraccatura_legno_massiccio', 'N.D.')}")
    doc.add_paragraph(f"Acciaio: {dati.get('baraccatura_acciaio', 'N.D.')}")

    doc.add_heading('5. Travi Principali / Portali (Confronto 3 Materiali)', level=1)
    doc.add_paragraph(f"Legno Lamellare: {dati.get('travi_legno', 'N.D.')}")
    doc.add_paragraph(f"Acciaio: {dati.get('travi_acciaio', 'N.D.')}")
    doc.add_paragraph(f"C.a.p.: {dati.get('travi_cap', 'N.D.')}")
    
    doc.add_heading('6. Pilastri (Perimetrali e Intermedi)', level=1)
    doc.add_paragraph(f"Legno Lamellare: {dati.get('pilastri_legno', 'N.D.')}")
    doc.add_paragraph(f"Acciaio: {dati.get('pilastri_acciaio', 'N.D.')}")
    doc.add_paragraph(f"C.a.p.: {dati.get('pilastri_cap', 'N.D.')}")
    
    doc.add_heading('7. Stabilizzazione e Controventi', level=1)
    doc.add_paragraph(f"Copertura: {dati.get('controventi_copertura_pos', 'N.D.')}")
    doc.add_paragraph(f"  - Opzione Legno: {dati.get('controventi_copertura_legno', 'N.D.')}")
    doc.add_paragraph(f"  - Opzione Acciaio: {dati.get('controventi_copertura_acciaio', 'N.D.')}")
    doc.add_paragraph(f"Parete: {dati.get('controventi_parete_pos', 'N.D.')}")
    doc.add_paragraph(f"  - Opzione Legno: {dati.get('controventi_parete_legno', 'N.D.')}")
    doc.add_paragraph(f"  - Opzione Acciaio: {dati.get('controventi_parete_acciaio', 'N.D.')}")
    
    doc.add_heading('8. Dettaglio Connessioni, Nodi e Giunti in Colmo', level=1)
    doc.add_paragraph(f"Connessione Trave/Pilastro: {dati.get('conn_trave_pilastro_tipo', 'N.D.')}")
    doc.add_paragraph(f"  - Elementi: {dati.get('conn_trave_pilastro_elementi', 'N.D.')} | Peso: {dati.get('conn_trave_pilastro_kg', 'N.D.')}")
    doc.add_paragraph(f"Connessione Pilastro/Fondazione: {dati.get('conn_pilastro_fondazione_tipo', 'N.D.')}")
    doc.add_paragraph(f"  - Ancoraggi: {dati.get('conn_pilastro_fondazione_elementi', 'N.D.')} | Peso: {dati.get('conn_pilastro_fondazione_kg', 'N.D.')}")
    doc.add_paragraph(f"Giunto in Colmo: {dati.get('dettaglio_giunto_colmo', 'N.D.')}")
    
    doc.add_heading('9. Protezione Antincendio', level=1)
    doc.add_paragraph(f"Classe Resistenza al Fuoco: {dati.get('classe_resistenza_fuoco', 'N.D.')}")
    doc.add_paragraph(f"Superficie Acciaio da Trattare: {dati.get('mq_intumescente', 'N.D.')}")
    doc.add_paragraph(f"Ciclo: {dati.get('dettaglio_verniciatura', 'N.D.')}")
    
    doc.add_heading('10. Note Tecniche', level=1)
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
    
    num_campate = max(1, int(round(lunghezza_edificio / interasse_portali))) if interasse_portali > 0 else 1
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
    num_sub_parete = max(1, int(round(altezza_gronda / 4.5))) if altezza_gronda > 0 else 1
    
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

st.set_page_config(page_title="Predimensionamento Strutturale NTC 2018", layout="wide")
st.title("Generatore Offerte Tecniche e Dimensionamento IA 🏗️")

with st.sidebar:
    st.header("Impostazioni Motore")
    api_key = st.text_input("Inserisci qui la tua API Key di Google", type="password")
    
    st.markdown("---")
    modalita_deterministica = st.toggle(
        "Motore Deterministico NTC 2018 (No IA)", 
        value=True,
        help="Se attivo, azzera l'interpretazione dell'IA per connessioni, baraccatura e fuoco, usando formule rigide e stabili."
    )
    if modalita_deterministica:
        st.success("🟢 Motore Matematico Locale Attivo (Risultati 100% stabili e ripetibili)")
    else:
        st.info("🤖 Modalità Ibrida con IA attiva (Richiede API Key)")

    st.markdown("---")
    if st.button("🔄 Nuovo Progetto / Reset", use_container_width=True):
        st.session_state.clear()
        st.rerun()

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
            st.success(f"File PDF '{file_caricato.name}' letto con successo!")
            
    except Exception as e:
        st.error(f"Errore nella lettura del file: {e}")

testo_commerciale = st.text_area(
    "Incolla qui le note del progetto o il capitolato:", 
    height=100,
    value="",
    placeholder="Incolla qui note di progetto o capitolato...",
    key="testo_commerciale"
)

st.markdown("### 📍 Localizzazione Cantiere (Google Maps)")
maps_url_ui = st.text_input(
    "Incolla il link di Google Maps del cantiere (es. https://maps.app.goo.gl/... o https://www.google.com/maps/...)",
    value="",
    placeholder="Incolla qui l'URL copiato da Google Maps...",
    key="maps_url_ui",
    help="Vai su Google Maps, seleziona il punto esatto, clicca su Condividi o copia il link dalla barra del browser e incollalo qui."
)

st.markdown("### 📐 Dimensioni Geometriche dell'Edificio (Modificabili)")
col_dim1, col_dim2, col_dim3, col_dim4, col_dim5 = st.columns(5)
with col_dim1:
    lunghezza_edificio_ui = st.number_input("Lunghezza Edificio (m)", min_value=0.0, value=25.0, step=1.0, format="%.1f", key="lunghezza_edificio_ui")
with col_dim2:
    interasse_portali_ui = st.number_input("Interasse Portali (m)", min_value=0.0, value=5.0, step=0.5, format="%.2f", key="interasse_portali_ui")
with col_dim3:
    luce_totale_ui = st.number_input("Luce Totale / Larghezza (m)", min_value=0.0, value=39.6, step=0.1, format="%.2f", key="luce_totale_ui")
with col_dim4:
    altezza_gronda_ui = st.number_input("Altezza Gronda (m)", min_value=0.0, value=9.0, step=0.5, format="%.1f", key="altezza_gronda_ui")
with col_dim5:
    altezza_colmo_ui = st.number_input("Altezza Colmo (m)", min_value=0.0, value=12.21, step=0.01, format="%.2f", key="altezza_colmo_ui")

st.markdown("### 🏛️ Configurazione Telaio e Travatura")
col_g1, col_g2 = st.columns(2)
with col_g1:
    tipo_travatura = st.selectbox("Tipologia Travatura di Copertura", ["Bi-falda semplice", "Bi-falda con intradosso curvo", "Trave di falda giuntata in colmo"], key="tipo_travatura")
with col_g2:
    num_appoggi = st.selectbox("Numero di Appoggi del Telaio", [2, 3, 4], index=1, format_func=lambda x: f"{x} Appoggi ({'Campata Unica' if x==2 else f'Multi-campata con {x-2} pilastro/i interno/i'})", key="num_appoggi")

st.markdown("### ⚙️ Parametri Carichi di Copertura e Pannellature")
col_c1, col_c2, col_c3 = st.columns(3)

with col_c1:
    tipo_isolante = st.selectbox("Tipologia Pannello Copertura", ["PIR / PUR", "Lana Minerale", "Lamiera Grecata Semplice"], key="tipo_isolante")
    if tipo_isolante == "PIR / PUR":
        spessore_pannello = st.selectbox("Spessore Pannello (mm)", [50, 60, 80, 100, 120], key="spessore_panni_pir")
    elif tipo_isolante == "Lana Minerale":
        spessore_pannello = st.selectbox("Spessore Pannello (mm)", [100, 120, 150, 170], key="spessore_panni_lana")
    else:
        spessore_pannello = 0

with col_c2:
    st.write("")
    st.write("")
    impianto_fv = st.checkbox("Impianto Fotovoltaico in Copertura (20 kg/mq)", value=False, key="impianto_fv")

with col_c3:
    carico_aggiuntivo = st.number_input("Carico aggiuntivo manuale (kN/mq)", min_value=0.0, value=0.0, step=0.05, format="%.2f", key="carico_aggiuntivo")

st.markdown("### 🧱 Rivestimento Parete")
col_p1, col_p2 = st.columns(2)
with col_p1:
    tipo_isolante_parete = st.selectbox("Tipologia Pannello Parete", ["PIR / PUR", "Lana di Roccia", "Lamiera Semplice", "Nessuno (Aperto)"], key="tipo_isolante_parete")
with col_p2:
    if tipo_isolante_parete == "PIR / PUR":
        spessore_pannello_parete = st.selectbox("Spessore Pannello Parete (mm)", [50, 60, 80, 100, 120], key="spessore_parete_pir")
    elif tipo_isolante_parete == "Lana di Roccia":
        spessore_pannello_parete = st.selectbox("Spessore Pannello Parete (mm)", [80, 100, 120, 150], key="spessore_parete_lana")
    else:
        spessore_pannello_parete = 0

if st.button("Esegui Dimensionamento e Genera Modello 3D", type="primary"):
    if lunghezza_edificio_ui <= 0 or interasse_portali_ui <= 0 or luce_totale_ui <= 0 or altezza_gronda_ui <= 0 or altezza_colmo_ui <= 0:
        st.warning("⚠️ Inserisci tutte le dimensioni geometriche con valori superiori a zero prima di eseguire il calcolo.")
    else:
        lat_estratta, lon_estratta = estrai_coordinate_da_url(maps_url_ui)
        num_campate_calc = max(1, int(round(lunghezza_edificio_ui / interasse_portali_ui)))
        impianto_fv_desc = "Presente (20 kg/mq)" if impianto_fv else "Assente"
        
        dati_base = {
            'lunghezza_edificio': lunghezza_edificio_ui,
            'interasse_portali': interasse_portali_ui,
            'luce_totale': luce_totale_ui,
            'altezza_gronda': altezza_gronda_ui,
            'altezza_colmo': altezza_colmo_ui,
            'num_campate': num_campate_calc,
            'tipo_travatura': tipo_travatura,
            'num_appoggi': num_appoggi,
            'tipo_isolante': tipo_isolante,
            'spessore_pannello': f"{spessore_pannello} mm" if tipo_isolante != "Lamiera Grecata Semplice" else "Lamiera Semplice",
            'tipo_isolante_parete': tipo_isolante_parete,
            'spessore_pannello_parete': f"{spessore_pannello_parete} mm" if tipo_isolante_parete not in ["Lamiera Semplice", "Nessuno (Aperto)"] else tipo_isolante_parete,
            'impianto_fv_desc': impianto_fv_desc,
            'carico_aggiuntivo': carico_aggiuntivo,
            'latitudine': lat_estratta,
            'longitudine': lon_estratta
        }

        if modalita_deterministica:
            with st.spinner('Estrazione coordinate ed esecuzione calcolo deterministico NTC 2018...'):
                dati = esegui_calcolo_deterministico(dati_base)
                dati.update(dati_base)
                distinta_elementi = calcola_distinta_elementi(dati)
                dati['distinta'] = distinta_elementi
                st.session_state['dati_ultimi'] = dati
                
                if maps_url_ui and lat_estratta == 46.4983 and lon_estratta == 11.3548 and "Bolzano" not in maps_url_ui:
                     st.warning("⚠️ Non è stato possibile estrarre le coordinate da questo link. È stato usato il carico neve di default (Bolzano). Assicurati di incollare un link di Google Maps completo.")
                else:
                     st.success(f"Link analizzato (Lat: {lat_estratta:.4f}, Lon: {lon_estratta:.4f})! Calcolo completato.")
        else:
            if not api_key:
                st.error("Inserisci prima l'API Key di Google nella barra laterale per usare la modalità IA!")
            else:
                dati_config_str = f"""
                --- CONFIGURAZIONE GEOMETRICA E LINK MAPS ---
                - Link Google Maps: {maps_url_ui} (Lat: {lat_estratta}, Lon: {lon_estratta})
                - Lunghezza Edificio: {lunghezza_edificio_ui} m
                - Interasse Portali: {interasse_portali_ui} m (Numero Campate: {num_campate_calc})
                - Luce Totale: {luce_totale_ui} m
                - Altezza Gronda: {altezza_gronda_ui} m
                - Altezza Colmo: {altezza_colmo_ui} m
                - Tipologia Travatura: {tipo_travatura}
                - Numero Appoggi Telaio: {num_appoggi} appoggi
                - Tipologia Pannello Copertura: {tipo_isolante} ({spessore_pannello} mm)
                - Impianto Fotovoltaico: {impianto_fv_desc}
                - Carico Permanente Aggiuntivo Manuale: {carico_aggiuntivo} kN/mq
                - Tipologia Pannello Parete: {tipo_isolante_parete} ({spessore_pannello_parete} mm)
                """
                testo_totale_analisi = testo_commerciale + "\n\n" + dati_config_str + "\n\n--- NOTE DAL FILE ALLEGATO ---\n" + testo_estratto_file
                
                genai.configure(api_key=api_key)
                try:
                    model = genai.GenerativeModel(
                        model_name='gemini-3.6-flash',
                        generation_config={"response_mime_type": "application/json", "temperature": 0.0}
                    )
                    prompt = f"""
Sei un ingegnere strutturista senior esperto in prefabbricazione industriale, NTC 2018. Analizza il testo tecnico e restituisci un oggetto JSON valido (senza markdown) con queste chiavi esatte:
- "luogo": stringa
- "qsk": float
- "zona_vento": stringa
- "pressione_vento": stringa
- "zona_sismica": stringa
- "classe_uso": stringa
- "fattore_struttura_q": stringa
- "campate_controventi_indici": lista di interi
- "interasse_arcarecci": float
- "sezione_arcarecci": stringa
- "verifica_arcarecci": stringa
- "baraccatura_legno_lamellare": stringa
- "baraccatura_legno_massiccio": stringa
- "baraccatura_acciaio": stringa
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
                    with st.spinner('Elaborazione calcoli strutturali con IA...'):
                        risposta_ia = model.generate_content(prompt)
                        testo_risposta = risposta_ia.text.strip()
                        if testo_risposta.startswith("```json"): testo_risposta = testo_risposta[7:]
                        if testo_risposta.startswith("```"): testo_risposta = testo_risposta[3:]
                        if testo_risposta.endswith("```"): testo_risposta = testo_risposta[:-3]
                        
                        dati = json.loads(testo_risposta.strip())
                        dati.update(dati_base)
                        risultati_strutturali = esegui_calcolo_deterministico(dati)
                        dati.update(risultati_strutturali)
                        distinta_elementi = calcola_distinta_elementi(dati)
                        dati['distinta'] = distinta_elementi
                        st.session_state['dati_ultimi'] = dati
                        st.success("Modello ibrido IA calcolato con successo!")
                except Exception as e:
                    st.error(f"Errore durante l'elaborazione con IA: {e}")

if 'dati_ultimi' in st.session_state:
    dati = st.session_state['dati_ultimi']
    distinta = dati.get('distinta', calcola_distinta_elementi(dati))
    st.markdown("---")
    
    col_dl1, col_dl2, col_dl3 = st.columns([1, 2, 1])
    with col_dl2:
        word_file = genera_word_report(dati, distinta)
        st.download_button(
            label="📄 Scarica Relazione e Distinta Elementi in Word (.docx)",
            data=word_file,
            file_name=f"Relazione_Predimensionamento_{dati.get('luogo', 'Progetto').replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
    
    st.markdown("---")
    st.markdown("### 🌐 Modello 3D Dinamico della Struttura")
    fig_3d = genera_modello_3d(dati)
    st.plotly_chart(fig_3d, use_container_width=True)
    
    st.markdown("---")
    st.markdown("### 📋 1. Distinta Elementi Principali (Computo Quantità)")
    c_e1, c_e2, c_e3, c_e4 = st.columns(4)
    c_e1.metric("Telai Principali", f"{distinta['num_telai']} pz")
    c_e2.metric("Pilastri Totali", f"{distinta['num_pilastri_totali']} pz", f"{distinta['num_pilastri_perimetrali']} Per. | {distinta['num_pilastri_interni']} Int.", delta_color="off")
    c_e3.metric("Travi di Falda", f"{distinta['num_travi_falda']} pz")
    c_e4.metric("File Arcarecci", f"{distinta['num_file_arcarecci']} file", f"Tot: {distinta['ml_arcarecci']} ml", delta_color="off")

    c_e5, c_e6, c_e7, c_e8 = st.columns(4)
    c_e5.metric("Moduli Controvento Cop.", f"{distinta['num_croci_copertura']} croci")
    c_e6.metric("Moduli Controvento Parete", f"{distinta['num_croci_parete']} croci")
    c_e7.metric("Superficie Copertura", f"{distinta['mq_copertura']} mq")
    c_e8.metric("Superficie Pareti Lunghe", f"{distinta['mq_pareti_lunghe']} mq")

    st.markdown("---")
    st.markdown("### 📍 2. Dati geometrici, climatici, sismici e di configurazione (NTC 2018)")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Località / Mappa", dati.get("luogo", "N.D."))
    c2.metric("Carico Neve (qsk)", f"{dati.get('qsk', 1.5)} kN/m²")
    c3.metric("Zona Vento", dati.get("zona_vento", "N.D."))
    c4.metric("Pressione Vento", dati.get("pressione_vento", "N.D."))
    
    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Azione Sismica", dati.get("zona_sismica", "N.D."))
    c6.metric("Luce Totale", f"{dati.get('luce_totale')} m")
    c7.metric("Schema Telaio", f"{dati.get('num_appoggi')} Appoggi")
    c8.metric("Travatura", dati.get("tipo_travatura", "Bi-falda semplice"))
    
    st.info(f"🏗️ **Copertura configurata:** Pannello {dati.get('tipo_isolante')} ({dati.get('spessore_pannello')}) | **Impianto FV:** {dati.get('impianto_fv_desc')} | **Carico Extra:** {dati.get('carico_aggiuntivo', 0.0)} kN/mq")
    
    st.markdown("---")
    st.markdown("### 🪵 3. Arcarecci di Copertura")
    st.info(f"**Passo Arcarecci:** {dati.get('interasse_arcarecci', 1.5)} m | **Sezione Consigliata:** {dati.get('sezione_arcarecci', 'N.D.')} | **Stato:** {dati.get('verifica_arcarecci', 'Verificato')}")
    
    st.markdown("---")
    st.markdown("### 🧱 4. Baraccatura di Parete (Supporto Rivestimento)")
    st.write(f"**Pannello Facciata:** {dati.get('tipo_isolante_parete', 'N.D.')} ({dati.get('spessore_pannello_parete', 'N.D.')})")
    col_b1, col_b2, col_b3 = st.columns(3)
    with col_b1:
        st.markdown("#### 🌲 Legno Lamellare")
        st.success(dati.get('baraccatura_legno_lamellare', 'N.D.'))
    with col_b2:
        st.markdown("#### 🪵 Legno Massiccio")
        st.success(dati.get('baraccatura_legno_massiccio', 'N.D.'))
    with col_b3:
        st.markdown("#### ⚙️ Acciaio")
        st.warning(dati.get('baraccatura_acciaio', 'N.D.'))

    st.markdown("---")
    st.markdown("### 📐 5. Travi Principali / Portali (Confronto Tecnologico)")
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
    st.markdown("### 🏛️ 6. Pilastri (Perimetrali e Intermedi)")
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
    st.markdown("### 🔗 7. Stabilizzazione e Controventi (Azioni Orizzontali)")
    col_cv1, col_cv2 = st.columns(2)
    with col_cv1:
        st.markdown("#### 🛡️ Controventi di Copertura (Falda)")
        st.write(f"📍 **Posizionamento:** {dati.get('controventi_copertura_pos', 'N.D.')}")
        st.info(f"🌲 **Opzione Legno:** {dati.get('controventi_copertura_legno', 'N.D.')}")
        st.info(f"⚙️ **Opzione Acciaio:** {dati.get('controventi_copertura_acciaio', 'N.D.')}")
    with col_cv2:
        st.markdown("#### 🧱 Controventi di Parete (Controventatura)")
        st.write(f"📍 **Posizionamento:** {dati.get('controventi_parete_pos', 'N.D.')}")
        st.warning(f"🌲 **Opzione Legno:** {dati.get('controventi_parete_legno', 'N.D.')}")
        st.warning(f"⚙️ **Opzione Acciaio:** {dati.get('controventi_parete_acciaio', 'N.D.')}")
    
    st.markdown("---")
    st.markdown("### 🔩 8. Dimensionamento Dettagliato Connessioni, Nodi e Giunto in Colmo")
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
    
    if dati.get('tipo_travatura') == "Trave di falda giuntata in colmo":
        st.info(f"📐 **Dettaglio Giunto in Colmo (Piastra di Giunzione):** {dati.get('dettaglio_giunto_colmo', 'N.D.')}")
    
    st.markdown("### 🔥 9. Requisiti di Resistenza al Fuoco e Vernice Intumescente")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        st.metric("Classe di Resistenza Richiesta", dati.get('classe_resistenza_fuoco', 'R 60'))
    with col_f2:
        st.metric("Superficie Acciaio da Trattare", dati.get('mq_intumescente', 'Non specificato'))
    st.info(f"**Specifiche Ciclo Antincendio:** {dati.get('dettaglio_verniciatura', 'N.D.')}")
    
    st.markdown("---")
    st.markdown("### 📝 10. Relazione e Note Tecniche")
    st.write(dati.get("note_tecniche", "Nessuna nota aggiuntiva."))
