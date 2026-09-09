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
import math

# --- 1. FUNZIONE ROBUSTA PER ESTRARRE COORDINATE ---
def estrai_dati_da_url_maps(url):
    url = url.strip()
    lat_def, lon_def = 46.4983, 11.3548
    nome_luogo_estratto = ""
    
    if not url:
        return lat_def, lon_def, ""
    
    match_place = re.search(r'/place/([^/@]+)', url)
    if match_place:
        nome_luogo_estratto = match_place.group(1).replace('+', ' ').split(',')[0].strip()

    if any(domain in url for domain in ["goo.gl", "googleusercontent.com", "maps.app.goo.gl"]):
        try:
            headers = {"User-Agent": "Mozilla/5.0"}
            response = requests.get(url, allow_redirects=True, timeout=5, headers=headers)
            match_place_redir = re.search(r'/place/([^/@]+)', response.url)
            if match_place_redir and not nome_luogo_estratto:
                nome_luogo_estratto = match_place_redir.group(1).replace('+', ' ').split(',')[0].strip()
        except Exception:
            pass 
            
    pattern_coords = [
        r'@([-0-9.]+),([-0-9.]+)',
        r'[?&](?:q|ll|s_loc)=([-0-9.]+)[,%]2[F0]([-0-9.]+)',
        r'[?&](?:q|ll|s_loc)=([-0-9.]+),([-0-9.]+)',
        r'!3d([-0-9.]+)!4d([-0-9.]+)',
        r'/(?:place|search)/([-0-9.]+),([-0-9.]+)',
        r'([-0-9.]+)\s*,\s*([-0-9.]+)'
    ]
    
    for pattern in pattern_coords:
        match = re.search(pattern, url)
        if match:
            try:
                lat, lon = float(match.group(1)), float(match.group(2))
                if 35.0 <= lat <= 47.5 and 6.0 <= lon <= 19.0:
                    return lat, lon, nome_luogo_estratto
            except ValueError:
                continue
                
    return lat_def, lon_def, nome_luogo_estratto

# --- 2. MOTORE DI CALCOLO STRUTTURALE DETERMINISTICO NTC 2018 ---
def estrai_parametri_ntc_da_coordinate_e_comune(lat, lon, comune_input=""):
    comune_pulito = comune_input.strip().lower()
    
    if "pantelleria" in comune_pulito or (36.7 <= lat <= 37.0 and 11.8 <= lon <= 12.1):
        alt_stim, zona_n, qsk, zona_v, press_v, zona_s = 50.0, "Zona III (Isole)", 0.50, "Zona 4", 0.58, "Zona 4"
        luogo_str = f"Pantelleria (TP) [GPS: {lat:.4f}, {lon:.4f}]"
    elif "lampedusa" in comune_pulito or (35.4 <= lat <= 35.6 and 12.5 <= lon <= 12.7):
        alt_stim, zona_n, qsk, zona_v, press_v, zona_s = 20.0, "Zona III (Isole)", 0.50, "Zona 4", 0.60, "Zona 4"
        luogo_str = f"Lampedusa (AG) [GPS: {lat:.4f}, {lon:.4f}]"
    elif any(x in comune_pulito for x in ["sardegna", "cagliari", "sassari", "nuoro", "oristano"]) or (8.0 <= lon <= 10.0 and 38.8 <= lat <= 41.3):
        alt_stim, zona_n, qsk, zona_v, press_v, zona_s = 50.0, "Zona III (Sardegna)", 0.50, "Zona 4", 0.55, "Zona 4"
        luogo_str = f"{comune_input.capitalize() if comune_input else 'Sardegna'} [GPS: {lat:.4f}, {lon:.4f}]"
    elif lat > 45.8:
        alt_stim, zona_n, zona_v, zona_s = 650.0, "Zona I (Alpina)", "Zona 1", "Zona 3"
        qsk = round(1.39 * (1.0 + (alt_stim / 728.0) ** 2), 2)
        press_v = round(0.50 * (1.0 + alt_stim/1000.0), 2)
        luogo_str = f"{comune_input.capitalize() if comune_input else 'Nord Italia (Alpi)'}"
    elif lat > 44.5:
        alt_stim, zona_n, zona_v, press_v, zona_s = 50.0, "Zona II (Padana)", "Zona 3", 0.48, "Zona 2/3"
        qsk = round(0.85 * (1.0 + (alt_stim / 778.0) ** 2), 2)
        luogo_str = f"{comune_input.capitalize() if comune_input else 'Nord Italia (Pianura)'}"
    elif lat > 41.0:
        alt_stim, zona_n, zona_v, press_v, zona_s = 150.0, "Zona II (Centro)", "Zona 2", 0.52, "Zona 1/2"
        qsk = round(0.85 * (1.0 + (alt_stim / 778.0) ** 2), 2)
        luogo_str = f"{comune_input.capitalize() if comune_input else 'Centro Italia'}"
    else:
        alt_stim, zona_n, zona_v, press_v, zona_s = 50.0, "Zona III (Sud)", "Zona 3/4", 0.58, "Zona 2"
        qsk = round(0.50 * (1.0 + (alt_stim / 833.0) ** 2), 2)
        luogo_str = f"{comune_input.capitalize() if comune_input else 'Sud Italia'}"
        
    return luogo_str, qsk, zona_v, f"{press_v} kN/mq", zona_s, alt_stim

def esegui_calcolo_deterministico(dati_geo):
    luce_totale = max(0.1, dati_geo['luce_totale'])
    interasse = max(0.1, dati_geo['interasse_portali'])
    h_gronda = max(0.1, dati_geo['altezza_gronda'])
    h_colmo = max(h_gronda + 0.1, dati_geo['altezza_colmo'])
    lunghezza_edificio = max(0.1, dati_geo['lunghezza_edificio'])
    num_appoggi = dati_geo['num_appoggi']
    pos_arcarecci = dati_geo.get('posizione_arcarecci', 'Sopra i telai')
    categoria_struttura = dati_geo.get('categoria_struttura', 'Portali ad anima piena')
    tipo_travatura = dati_geo.get('tipo_travatura', 'Bi-falda semplice')
    
    luogo_str, qsk, zona_vento, press_vento_str, zona_sismica, alt_stimata = estrai_parametri_ntc_da_coordinate_e_comune(
        dati_geo.get('latitudine', 46.49), dati_geo.get('longitudine', 11.35), dati_geo.get('comune', '')
    )
    pressione_vento = float(press_vento_str.split()[0])
    
    sviluppo_falda = math.sqrt((luce_totale/2)**2 + (h_colmo - h_gronda)**2)
    spessore_cop = str(dati_geo.get('spessore_pannello', ''))
    tipo_cop = str(dati_geo.get('tipo_isolante', ''))
    
    if "Lamiera" in tipo_cop: max_passo_arc = 1.2
    elif "50" in spessore_cop or "60" in spessore_cop: max_passo_arc = 1.8
    elif "80" in spessore_cop or "100" in spessore_cop: max_passo_arc = 2.5
    else: max_passo_arc = 3.0
    
    num_campi_falda = math.ceil(sviluppo_falda / max_passo_arc)
    passo_arcarecci = sviluppo_falda / num_campi_falda if num_campi_falda > 0 else max_passo_arc

    spessore_par = str(dati_geo.get('spessore_pannello_parete', ''))
    tipo_par = str(dati_geo.get('tipo_isolante_parete', ''))
    
    if "Lamiera" in tipo_par or "Nessuno" in tipo_par: max_passo_bar = 1.5
    elif "50" in spessore_par or "60" in spessore_par: max_passo_bar = 2.2
    elif "80" in spessore_par or "100" in spessore_par: max_passo_bar = 3.0
    else: max_passo_bar = 3.5
    
    num_campi_parete = math.ceil(h_gronda / max_passo_bar)
    passo_baraccatura = h_gronda / num_campi_parete if num_campi_parete > 0 else max_passo_bar

    ml_baraccatura_long_singola = int(h_gronda / passo_baraccatura) * lunghezza_edificio
    ml_baraccatura_long_tot = ml_baraccatura_long_singola * 2
    
    ml_baraccatura_timpani_singolo = 0
    z = passo_baraccatura
    while z < h_colmo:
        if z <= h_gronda: larghezza_timpano = luce_totale
        else: larghezza_timpano = luce_totale * (1 - ((z - h_gronda) / (h_colmo - h_gronda)))
        ml_baraccatura_timpani_singolo += larghezza_timpano
        z += passo_baraccatura
    ml_baraccatura_timpani_tot = ml_baraccatura_timpani_singolo * 2
    ml_tot_baraccatura = ml_baraccatura_long_tot + ml_baraccatura_timpani_tot

    g2 = 0.25 + 0.20 if "Presente" in dati_geo.get('impianto_fv_desc', '') else 0.25
    g2 += dati_geo.get('carico_aggiuntivo', 0.0)
    q_tot_copertura_mq = (1.3 * 0.15 + 1.5 * g2 + 1.5 * qsk)
    
    M_ed_arc = (q_tot_copertura_mq * passo_arcarecci * interasse**2) / (8 if "In luce" in pos_arcarecci else 10)
    W_req_arc = (M_ed_arc * 100) / 27.5
    if W_req_arc < 30: sez_arc = "Tubolare 100x50x3"
    elif W_req_arc < 55: sez_arc = "Tubolare 120x60x4"
    elif W_req_arc < 85: sez_arc = "Tubolare 150x75x4 / IPE 140"
    else: sez_arc = "IPE 180"

    luce_campata = luce_totale / (num_appoggi - 1) if num_appoggi >= 3 else luce_totale
    q_ed_portale = interasse * q_tot_copertura_mq
    m_ed = (q_ed_portale * (luce_campata ** 2)) / (10.0 if num_appoggi >= 3 else 8.0)
    v_ed = (q_ed_portale * luce_campata) / 2.0
    H_truss = max(0.5, h_colmo - h_gronda)

    peso_stimato_trave_kg_ml = 45.0
    
    if categoria_struttura == "Capriate":
        travi_legno_out = f"Catena Inf. {luce_totale}m (24x28 cm) | Puntoni (24x28 cm)"
        travi_acciaio_out = f"Catena Inf. Tubolare 120x120x5 | Puntoni Tub. 120x120x5"
        peso_stimato_trave_kg_ml = 35.0
    elif categoria_struttura == "Travi Reticolari":
        travi_legno_out = f"Correnti L={luce_totale}m (GL24h 24x32 cm) | Diag. 16x16 cm"
        travi_acciaio_out = f"Correnti L={luce_totale}m (Tub. 150x150x6) | Diag. 100x100x5"
        peso_stimato_trave_kg_ml = 45.0
    else:
        w_req_cm3 = (m_ed * 1e6) / 14500.0  
        h_legno_cm = max(44, int((6 * w_req_cm3 / 20) ** 0.5))
        travi_legno_out = f"Base 20 cm x Altezza {h_legno_cm} cm (GL24h)"
        
        w_el_req_cm3 = (m_ed * 100.0) / 33.8 
        if w_el_req_cm3 > 3500: prof_acc = "IPE 600 / HEB 500"
        elif w_el_req_cm3 > 2000: prof_acc = "IPE 500 / HEA 400"
        elif w_el_req_cm3 > 1000: prof_acc = "IPE 400 / HEA 300"
        else: prof_acc = "IPE 330 / HEA 240"
        travi_acciaio_out = f"Profilo {prof_acc} S355JR"
        
        peso_stimato_trave_kg_ml = (20 * h_legno_cm / 10000) * 500

    M_base_vento = (pressione_vento * interasse * h_gronda**2) / 2
    h_pil_perim_cm = max(32, int(min(140, 32 + M_base_vento * 0.5)))
    peso_stimato_pilastro_kg_ml = (20 * h_pil_perim_cm / 10000) * 500 

    n_bulloni = max(6, int(v_ed / 25.0) * 2)
    peso_conn = round(n_bulloni * 4.0 + 25.0, 1)

    passo_max_montanti = 6.0 
    num_sottocampate_timpano = max(1, int(np.ceil(luce_campata / passo_max_montanti)))
    num_montanti_timpano_singola_facciata = (num_sottocampate_timpano - 1) * (num_appoggi - 1)
    passo_montanti_timpano = luce_campata / num_sottocampate_timpano if num_sottocampate_timpano > 0 else 0

    return {
        "luogo": luogo_str, "qsk": qsk, "zona_vento": zona_vento, "pressione_vento": press_vento_str, "zona_sismica": zona_sismica,
        "travi_legno": travi_legno_out, "travi_acciaio": travi_acciaio_out, "travi_cap": "C.A.P. da verificare a parte",
        "pilastri_perimetrali_legno": f"Sezione 20x{h_pil_perim_cm} cm", "pilastri_perimetrali_acciaio": "Profilo HEA 200",
        "passo_arcarecci_calc": round(passo_arcarecci, 2), "sezione_arcarecci": sez_arc,
        "passo_baraccatura_calc": round(passo_baraccatura, 2),
        "ml_baraccatura_tot": round(ml_tot_baraccatura, 1),
        "peso_stimato_trave_kg_ml": peso_stimato_trave_kg_ml,
        "peso_stimato_pilastro_kg_ml": peso_stimato_pilastro_kg_ml,
        "num_montanti_timpano_singolo": num_montanti_timpano_singola_facciata,
        "passo_montanti_timpano": round(passo_montanti_timpano, 2),
        "conn_trave_pilastro_perim_elementi": f"N. {n_bulloni} bulloni 8.8 M20",
        "conn_trave_pilastro_perim_kg": f"{peso_conn} kg",
        "mq_intumescente": round((luce_totale + h_gronda * 2) * (dati_geo['num_campate'] + 1) * 0.6, 1)
    }

# --- 3. DISTINTA ELEMENTI ---
def calcola_distinta_elementi(dati):
    L, B, i_portali = dati['lunghezza_edificio'], dati['luce_totale'], dati['interasse_portali']
    num_campate = max(1, int(round(L / i_portali)))
    num_telai = num_campate + 1
    num_pilastri_totali = num_telai * dati['num_appoggi']
    
    sviluppo_falda = ((B/2)**2 + (dati['altezza_colmo'] - dati['altezza_gronda'])**2)**0.5
    mq_copertura = L * sviluppo_falda * 2
    mq_pareti = L * dati['altezza_gronda'] * 2 
    mq_timpani = 2 * (B * dati['altezza_gronda'] + (B * (dati['altezza_colmo'] - dati['altezza_gronda']) / 2))

    return {
        "num_telai": num_telai,
        "num_pilastri_totali": num_pilastri_totali,
        "num_pilastri_perimetrali": num_telai * 2,
        "num_pilastri_interni": num_pilastri_totali - (num_telai * 2),
        "num_travi_falda": num_telai * 2,
        "num_file_arcarecci": (math.ceil(sviluppo_falda / dati.get('passo_arcarecci_calc', 1.5)) * 2) - 1,
        "ml_arcarecci": round(((math.ceil(sviluppo_falda / dati.get('passo_arcarecci_calc', 1.5)) * 2) - 1) * L, 1),
        "mq_copertura": round(mq_copertura, 1),
        "mq_pareti_lunghe": round(mq_pareti, 1),
        "mq_timpani": round(mq_timpani, 1)
    }

# --- 4. LOGISTICA AVANZATA (INCROCIO PESO/VOLUME/PORTATA) ---
def calcola_logistica_trasporti(dati, distinta):
    sviluppo_falda = math.sqrt((dati['luce_totale']/2)**2 + (dati['altezza_colmo'] - dati['altezza_gronda'])**2)
    cat_strut = dati.get('categoria_struttura', 'Portali ad anima piena')
    
    if cat_strut == "Portali ad anima piena": max_lunghezza_trave = sviluppo_falda
    else: max_lunghezza_trave = min(12.0, sviluppo_falda)

    portata_max_kg = 24000.0 # Portata operativa utile media per evitare sforamenti
    peso_totale_singola_trave = max_lunghezza_trave * dati.get('peso_stimato_trave_kg_ml', 45.0)
    
    if max_lunghezza_trave <= 13.6: mezzo_travi, limite_qta_ingombro = "Bilico Standard 13.6m", 8
    elif max_lunghezza_trave <= 16.0: mezzo_travi, limite_qta_ingombro = "Eccezionale Allungato 16m", 6
    elif max_lunghezza_trave <= 20.0: mezzo_travi, limite_qta_ingombro = "Eccezionale 20m", 4
    elif max_lunghezza_trave <= 25.0: mezzo_travi, limite_qta_ingombro = "Eccezionale Allungabile 25m", 2
    else: mezzo_travi, limite_qta_ingombro = "Eccezionale >25m (Scorta Tecnica)", 1

    limite_qta_peso = max(1, int(portata_max_kg / max(1.0, peso_totale_singola_trave)))
    qta_effettiva_per_viaggio = min(limite_qta_ingombro, limite_qta_peso)
    viaggi_travi = math.ceil(distinta['num_travi_falda'] / qta_effettiva_per_viaggio)

    max_h_pil = max(dati['altezza_gronda'], dati['altezza_colmo'])
    peso_singolo_pilastro = max_h_pil * dati.get('peso_stimato_pilastro_kg_ml', 60.0)
    mezzo_pilastri = "Bilico Standard 13.6m" if max_h_pil <= 13.6 else "Eccezionale Allungato"
    qta_pil_ingombro = 12 if max_h_pil <= 13.6 else 6
    qta_pil_peso = max(1, int(portata_max_kg / max(1.0, peso_singolo_pilastro)))
    
    viaggi_pilastri = math.ceil(distinta['num_pilastri_totali'] / min(qta_pil_ingombro, qta_pil_peso))

    ml_tot_profili = distinta['ml_arcarecci'] + dati.get('ml_baraccatura_tot', 0)
    peso_profili = ml_tot_profili * 8.5 
    viaggi_profili = max(1, math.ceil(peso_profili / portata_max_kg))
    
    mq_tot_rivestimenti = distinta['mq_copertura'] + distinta['mq_pareti_lunghe'] + distinta['mq_timpani']
    viaggi_pannelli = max(1, math.ceil(mq_tot_rivestimenti / 550.0))
    viaggi_accessori = 1

    return {
        "max_lunghezza_trave": round(max_lunghezza_trave, 2),
        "mezzo_travi": mezzo_travi, "viaggi_travi": viaggi_travi,
        "qta_effettiva_travi": qta_effettiva_per_viaggio,
        "mezzo_pilastri": mezzo_pilastri, "viaggi_pilastri": viaggi_pilastri,
        "ml_tot_profili": round(ml_tot_profili, 1),
        "viaggi_profili": viaggi_profili,
        "mq_tot_rivestimenti": round(mq_tot_rivestimenti, 1),
        "viaggi_pannelli": viaggi_pannelli, "viaggi_accessori": viaggi_accessori,
        "tot_viaggi": viaggi_travi + viaggi_pilastri + viaggi_profili + viaggi_pannelli + viaggi_accessori
    }

# --- 5. EXPORT WORD (REPORT COMPLETO) ---
def genera_word_report(dati, distinta, logistica):
    doc = Document()
    doc.add_heading('Relazione Strutturale e Logistica NTC 2018', 0)
    
    doc.add_heading('1. Dati Cantiere e Geometria', level=1)
    doc.add_paragraph(f"Località: {dati.get('luogo', 'N.D.')}")
    doc.add_paragraph(f"Dimensioni: {dati.get('lunghezza_edificio')}x{dati.get('luce_totale')} H={dati.get('altezza_gronda')}m")
    
    doc.add_heading('2. Distinta Elementi', level=1)
    doc.add_paragraph(f"Telai: {distinta['num_telai']} | Pilastri: {distinta['num_pilastri_totali']}")
    doc.add_paragraph(f"Travi di falda: {distinta['num_travi_falda']}")
    doc.add_paragraph(f"Sviluppo Arcarecci: {distinta['ml_arcarecci']} ml")
    
    doc.add_heading('3. Logistica Trasporti (Flotta Veneta)', level=1)
    doc.add_paragraph(f"Totale Viaggi Stimati: {logistica['tot_viaggi']}")
    doc.add_paragraph(f"- Travi Principali: {logistica['viaggi_travi']} viaggi su {logistica['mezzo_travi']}")
    doc.add_paragraph(f"- Pilastri: {logistica['viaggi_pilastri']} viaggi su {logistica['mezzo_pilastri']}")
    doc.add_paragraph(f"- Arcarecci/Profili: {logistica['viaggi_profili']} viaggi standard")
    doc.add_paragraph(f"- Pannelli/Copertura: {logistica['viaggi_pannelli']} viaggi standard")
    
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# --- 6. MODELLO 3D (PLOTLY) ---
def genera_modello_3d(dati):
    fig = go.Figure()
    luce = dati.get('luce_totale', 39.6)
    h_g = dati.get('altezza_gronda', 9.0)
    h_c = dati.get('altezza_colmo', 12.21)
    L = dati.get('lunghezza_edificio', 25.0)
    i_p = dati.get('interasse_portali', 5.0)
    n_app = dati.get('num_appoggi', 2)
    cat_strut = dati.get('categoria_struttura', 'Portali ad anima piena')
    
    num_camp = max(1, int(round(L / i_p))) if i_p > 0 else 1
    y_portali = [i * i_p for i in range(num_camp + 1)]
    
    if n_app == 2: x_pil = [0.0, luce]
    elif n_app == 3: x_pil = [0.0, luce / 2.0, luce]
    else: x_pil = [0.0, luce / 3.0, (2 * luce) / 3.0, luce]
        
    for idx_y, y in enumerate(y_portali):
        for idx_x, x in enumerate(x_pil):
            hp = h_g if (x == 0.0 or x == luce) else h_c
            fig.add_trace(go.Scatter3d(x=[x, x], y=[y, y], z=[0, hp], mode='lines', line=dict(color='darkblue', width=6), showlegend=(idx_y==0 and idx_x==0), name='Pilastri'))
        
        show_trave = (idx_y == 0)
        if cat_strut == "Capriate":
            fig.add_trace(go.Scatter3d(x=[0, luce], y=[y, y], z=[h_g, h_g], mode='lines', line=dict(color='saddlebrown', width=6), showlegend=show_trave, name='Catena'))
            fig.add_trace(go.Scatter3d(x=[0, luce/2, luce], y=[y, y, y], z=[h_g, h_c, h_g], mode='lines', line=dict(color='firebrick', width=6), showlegend=show_trave, name='Puntoni'))
        elif cat_strut == "Travi Reticolari":
            fig.add_trace(go.Scatter3d(x=[0, luce], y=[y, y], z=[h_g, h_g], mode='lines', line=dict(color='gray', width=6), showlegend=show_trave, name='Corrente Inf'))
            fig.add_trace(go.Scatter3d(x=[0, luce/2, luce], y=[y, y, y], z=[h_g, h_c, h_g], mode='lines', line=dict(color='dimgray', width=6), showlegend=show_trave, name='Corrente Sup'))
        else:
            fig.add_trace(go.Scatter3d(x=[0, luce/2, luce], y=[y, y, y], z=[h_g, h_c, h_g], mode='lines', line=dict(color='firebrick', width=6), showlegend=show_trave, name='Travi di Falda'))

    fig.update_layout(
        title="Modello 3D Dinamico",
        scene=dict(xaxis_title='X (Larghezza)', yaxis_title='Y (Lunghezza)', zaxis_title='Z (Altezza)', aspectmode='data'),
        margin=dict(l=0, r=0, b=0, t=40), height=550
    )
    return fig

# --- 7. INTERFACCIA STREAMLIT COMPLETA ---
st.set_page_config(page_title="Predimensionamento IA NTC", layout="wide")
st.title("Generatore Offerte Tecniche, Logistica e Modello 3D 🏗️")

with st.sidebar:
    st.header("Impostazioni")
    api_key = st.text_input("API Key di Google (Opzionale per IA)", type="password")
    modalita_deterministica = st.toggle("Motore Matematico Locale Attivo", value=True)
    if st.button("🔄 Reset Progetto", use_container_width=True):
        st.session_state.clear()
        st.rerun()

st.markdown("### 📍 Localizzazione Cantiere (Google Maps e Comune)")
c_loc1, c_loc2 = st.columns([2, 1])
with c_loc1: url_maps = st.text_input("Link Google Maps:")
with c_loc2: comune = st.text_input("Comune di installazione:")

st.markdown("### 📐 Dimensioni Geometriche dell'Edificio")
col1, col2, col3, col4, col5 = st.columns(5)
with col1: L_ed = st.number_input("Lunghezza (m)", value=25.0, step=1.0)
with col2: i_port = st.number_input("Interasse (m)", value=5.0, step=0.5)
with col3: B_luce = st.number_input("Luce Totale (m)", value=39.6, step=0.1)
with col4: H_gronda = st.number_input("H Gronda (m)", value=9.0, step=0.5)
with col5: H_colmo = st.number_input("H Colmo (m)", value=12.21, step=0.1)

st.markdown("### 🏛️ Configurazione Telaio e Travatura")
c_g1, c_g2, c_g3 = st.columns(3)
with c_g1: cat_strutt = st.selectbox("Categoria Struttura", ["Portali ad anima piena", "Capriate", "Travi Reticolari"])
with c_g2: num_appoggi = st.selectbox("Numero Appoggi Telaio", [2, 3, 4], index=1)
with c_g3: pos_arc = st.radio("Posizionamento Arcarecci", ["Sopra i telai (Continuo)", "In luce (Semplice appoggio)"])

st.markdown("### ⚙️ Parametri Carichi e Pannellature")
c_c1, c_c2, c_c3 = st.columns(3)
with c_c1:
    tipo_isol = st.selectbox("Copertura", ["PIR / PUR", "Lana Minerale", "Lamiera Grecata Semplice"])
    spess_isol = st.selectbox("Spessore Copertura (mm)", [50, 60, 80, 100, 120])
with c_c2:
    tipo_par = st.selectbox("Parete", ["PIR / PUR", "Lana di Roccia", "Nessuno (Aperto)"])
    spess_par = st.selectbox("Spessore Parete (mm)", [50, 60, 80, 100, 120])
with c_c3:
    impianto_fv = st.checkbox("Impianto FV (+20 kg/mq)", value=False)
    carico_agg = st.number_input("Carico extra (kN/mq)", value=0.0, step=0.1)

if st.button("🚀 Esegui Calcolo Strutturale, Logistica e 3D", type="primary"):
    if L_ed > 0 and i_port > 0 and B_luce > 0:
        lat, lon, place = estrai_dati_da_url_maps(url_maps)
        com_fin = comune if comune else place
        
        dati_base = {
            'lunghezza_edificio': L_ed, 'interasse_portali': i_port, 'luce_totale': B_luce,
            'altezza_gronda': H_gronda, 'altezza_colmo': H_colmo, 'num_campate': max(1, int(L_ed/i_port)),
            'categoria_struttura': cat_strutt, 'tipo_travatura': 'Standard', 'num_appoggi': num_appoggi,
            'posizione_arcarecci': pos_arc, 'tipo_isolante': tipo_isol, 'spessore_pannello': spess_isol,
            'tipo_isolante_parete': tipo_par, 'spessore_pannello_parete': spess_par,
            'impianto_fv_desc': "Presente" if impianto_fv else "Assente", 'carico_aggiuntivo': carico_agg,
            'latitudine': lat, 'longitudine': lon, 'comune': com_fin
        }
        
        with st.spinner("Calcolo in corso..."):
            dati = esegui_calcolo_deterministico(dati_base)
            dati.update(dati_base)
            dati['distinta'] = calcola_distinta_elementi(dati)
            dati['logistica'] = calcola_logistica_trasporti(dati, dati['distinta'])
            st.session_state['risultati'] = dati
            st.success("Analisi e Generazione Modello completata!")
    else:
        st.error("Inserisci parametri geometrici validi (>0).")

# --- 8. OUTPUT A SCHEDE (TABS) ---
if 'risultati' in st.session_state:
    res = st.session_state['risultati']
    dist = res['distinta']
    logi = res['logistica']
    
    st.markdown("---")
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Computo & Struttura (NTC)", "🚚 Piano Logistico Avanzato", "🌐 Modello 3D", "📄 Relazione Word"])
    
    with tab1:
        st.subheader("Dati NTC 2018 e Computo Elementi")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Località NTC", res['luogo'].split('[')[0])
        c2.metric("Qsk Neve", f"{res['qsk']} kN/m²")
        c3.metric("Telai Totali", f"{dist['num_telai']} pz")
        c4.metric("Travi Principali", f"{dist['num_travi_falda']} pz")
        
        st.markdown("#### Dimensionamento Elementi")
        st.info(f"**Predimensionamento Legno:** {res['travi_legno']}")
        st.warning(f"**Predimensionamento Acciaio:** {res['travi_acciaio']}")
        st.write(f"- **Passo Arcarecci:** {res['passo_arcarecci_calc']} m (Sez: {res['sezione_arcarecci']})")
        st.write(f"- **Passo Baraccatura Parete:** {res['passo_baraccatura_calc']} m (Totale ML: {res['ml_baraccatura_tot']})")
        st.success(f"**Nodi Perimetrali:** {res['conn_trave_pilastro_perim_elementi']} (stima acciaio {res['conn_trave_pilastro_perim_kg']})")

    with tab2:
        st.subheader("Incrocio Limiti Dimensionali e Ponderali (Flotta Veneta)")
        st.metric("Totale Viaggi Stimati", f"{logi['tot_viaggi']} Viaggi")
        
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            st.markdown("#### 🏗️ Trasporto Travi di Falda")
            st.write(f"- **Lunghezza singolo pezzo:** {logi['max_lunghezza_trave']} m")
            st.write(f"- **Mezzo selezionato:** {logi['mezzo_travi']}")
            st.write(f"- **Pezzi max per viaggio (Saturazione Peso/Volume):** {logi['qta_effettiva_travi']} pz")
            st.success(f"**Totale Viaggi Travi:** {logi['viaggi_travi']}")
            
        with col_t2:
            st.markdown("#### 🏛️ Trasporto Pilastri e Accessori")
            st.write(f"- **Mezzo Pilastri:** {logi['mezzo_pilastri']}")
            st.success(f"**Totale Viaggi Pilastri:** {logi['viaggi_pilastri']}")
            st.info(f"**Viaggi Arcarecci/Baraccature:** {logi['viaggi_profili']} (Bilico Standard - {logi['ml_tot_profili']} ml)")
            st.info(f"**Viaggi Pannelli:** {logi['viaggi_pannelli']} (Bilico Standard - {logi['mq_tot_rivestimenti']} mq volume)")

    with tab3:
        st.subheader("Visualizzazione Strutturale Interattiva")
        fig_3d = genera_modello_3d(res)
        st.plotly_chart(fig_3d, use_container_width=True)

    with tab4:
        st.subheader("Esportazione Documentazione")
        word_file = genera_word_report(res, dist, logi)
        st.download_button(
            label="📄 Scarica Relazione Tecnica e Logistica (.docx)",
            data=word_file,
            file_name=f"Relazione_{res.get('luogo', 'Progetto').split(' ')[0]}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )
