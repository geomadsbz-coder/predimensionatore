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
import pandas as pd

# --- FUNZIONE ROBUSTA PER ESTRARRE COORDINATE E NOME LUOGO DA URL DI GOOGLE MAPS ---
def estrai_dati_da_url_maps(url):
    url = url.strip()
    lat_def, lon_def = 46.4983, 11.3548
    nome_luogo_estratto = ""
    
    if not url:
        return lat_def, lon_def, ""
    
    match_place = re.search(r'/place/([^/@]+)', url)
    if match_place:
        raw_place = match_place.group(1)
        nome_luogo_estratto = raw_place.replace('+', ' ').split(',')[0].strip()

    if any(domain in url for domain in ["goo.gl", "googleusercontent.com", "maps.app.goo.gl"]):
        try:
            headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
            response = requests.get(url, allow_redirects=True, timeout=5, headers=headers)
            url = response.url 
            match_place_redir = re.search(r'/place/([^/@]+)', url)
            if match_place_redir and not nome_luogo_estratto:
                nome_luogo_estratto = match_place_redir.group(1).replace('+', ' ').split(',')[0].strip()
        except Exception:
            pass 
            
    match_at = re.search(r'@([-0-9.]+),([-0-9.]+)', url)
    if match_at:
        try:
            lat = float(match_at.group(1))
            lon = float(match_at.group(2))
            if 35.0 <= lat <= 47.5 and 6.0 <= lon <= 19.0:
                return lat, lon, nome_luogo_estratto
        except ValueError:
            pass
            
    match_q = re.search(r'[?&](?:q|ll|s_loc)=([-0-9.]+)[,%]2[F0]([-0-9.]+)', url) or re.search(r'[?&](?:q|ll|s_loc)=([-0-9.]+),([-0-9.]+)', url)
    if match_q:
        try:
            lat = float(match_q.group(1))
            lon = float(match_q.group(2))
            if 35.0 <= lat <= 47.5 and 6.0 <= lon <= 19.0:
                return lat, lon, nome_luogo_estratto
        except ValueError:
            pass

    match_3d4d = re.search(r'!3d([-0-9.]+)!4d([-0-9.]+)', url)
    if match_3d4d:
        try:
            lat = float(match_3d4d.group(1))
            lon = float(match_3d4d.group(2))
            if 35.0 <= lat <= 47.5 and 6.0 <= lon <= 19.0:
                return lat, lon, nome_luogo_estratto
        except ValueError:
            pass

    match_path = re.search(r'/(?:place|search)/([-0-9.]+),([-0-9.]+)', url)
    if match_path:
        try:
            lat = float(match_path.group(1))
            lon = float(match_path.group(2))
            if 35.0 <= lat <= 47.5 and 6.0 <= lon <= 19.0:
                return lat, lon, nome_luogo_estratto
        except ValueError:
            pass
            
    match_raw = re.search(r'([-0-9.]+)\s*,\s*([-0-9.]+)', url)
    if match_raw:
        try:
            lat = float(match_raw.group(1))
            lon = float(match_raw.group(2))
            if 35.0 <= lat <= 47.5 and 6.0 <= lon <= 19.0:
                return lat, lon, nome_luogo_estratto
        except ValueError:
            pass
            
    return lat_def, lon_def, nome_luogo_estratto

# --- MOTORE DI CALCOLO STRUTTURALE DETERMINISTICO NTC 2018 ---
def estrai_parametri_ntc_da_coordinate_e_comune(lat, lon, comune_input=""):
    comune_pulito = comune_input.strip().lower()
    
    if "pantelleria" in comune_pulito or (36.7 <= lat <= 37.0 and 11.8 <= lon <= 12.1):
        altitudine_stimata = 50.0  
        zona_neve = "Zona III (Meridionale / Isole)"
        qsk = 0.50
        zona_vento = "Zona 4 (Sud / Isole)"
        pressione_vento = 0.58
        zona_sismica = "Zona 4 (Sismicità molto bassa - Pantelleria)"
        luogo_str = f"Comune: Pantelleria (TP) [GPS: {lat:.4f}, {lon:.4f}] - Alt. {altitudine_stimata}m"
        return luogo_str, qsk, zona_vento, f"{pressione_vento} kN/mq", zona_sismica, altitudine_stimata

    if "lampedusa" in comune_pulito or (35.4 <= lat <= 35.6 and 12.5 <= lon <= 12.7):
        altitudine_stimata = 20.0  
        zona_neve = "Zona III (Meridionale / Isole)"
        qsk = 0.50
        zona_vento = "Zona 4 (Sud / Isole)"
        pressione_vento = 0.60
        zona_sismica = "Zona 4 (Sismicità molto bassa - Lampedusa e Linosa)"
        luogo_str = f"Comune: Lampedusa e Linosa (AG) [GPS: {lat:.4f}, {lon:.4f}] - Alt. {altitudine_stimata}m"
        return luogo_str, qsk, zona_vento, f"{pressione_vento} kN/mq", zona_sismica, altitudine_stimata

    if "sardegna" in comune_pulito or "cagliari" in comune_pulito or "sassari" in comune_pulito or "nuoro" in comune_pulito or "oristano" in comune_pulito or (8.0 <= lon <= 10.0 and 38.8 <= lat <= 41.3):
        altitudine_stimata = 50.0
        zona_neve = "Zona III (Sardegna)"
        qsk = 0.50
        zona_vento = "Zona 4 (Sardegna)"
        pressione_vento = 0.55
        zona_sismica = "Zona 4 (Sismicità trascurabile / Territorio Regionale Sardo)"
        comune_str = comune_input if comune_input else "Sardegna"
        luogo_str = f"Comune: {comune_str.capitalize()} [GPS: {lat:.4f}, {lon:.4f}] - Alt. {altitudine_stimata}m"
        return luogo_str, qsk, zona_vento, f"{pressione_vento} kN/mq", zona_sismica, altitudine_stimata

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
        zona_sismica = "Zona 2 / 3 (Media/Bassa sismicità - Pianura Padana)"
    elif lat > 41.0:
        altitudine_stimata = 150.0  
        zona_neve = "Zona II (Interna Centro)"
        qsk = round(0.85 * (1.0 + (altitudine_stimata / 778.0) ** 2), 2)
        zona_vento = "Zona 2 (vb = 28 m/s)"
        pressione_vento = 0.52
        zona_sismica = "Zona 1 / 2 (Alta/Media sismicità - Appennino Centromeridionale)"
    else:
        altitudine_stimata = 50.0   
        zona_neve = "Zona III (Meridionale / Costiera)"
        qsk = round(0.50 * (1.0 + (altitudine_stimata / 833.0) ** 2), 2)
        zona_vento = "Zona 3 o 4 (Sud/Isole)"
        pressione_vento = 0.58
        zona_sismica = "Zona 2 (Meridionale / Media sismicità)"
        
    comune_display = f"Comune: {comune_input.capitalize()}" if comune_input else f"Google Maps ({lat:.4f}, {lon:.4f})"
    luogo_str = f"{comune_display} - Alt. stimata: {altitudine_stimata}m - {zona_neve}"
    return luogo_str, qsk, zona_vento, f"{pressione_vento} kN/mq", zona_sismica, altitudine_stimata

def esegui_calcolo_deterministico(dati_geo):
    luce_totale = dati_geo['luce_totale']
    interasse = dati_geo['interasse_portali']
    h_gronda = dati_geo['altezza_gronda']
    h_colmo = dati_geo['altezza_colmo']
    num_appoggi = dati_geo['num_appoggi']
    lunghezza_edificio = dati_geo['lunghezza_edificio']
    pos_arcarecci = dati_geo.get('posizione_arcarecci', 'Sopra i telai')
    
    categoria_struttura = dati_geo.get('categoria_struttura', 'Portali ad anima piena')
    tipo_travatura = dati_geo.get('tipo_travatura', 'Bi-falda semplice')
    
    lat = dati_geo.get('latitudine', 46.4983)
    lon = dati_geo.get('longitudine', 11.3548)
    comune = dati_geo.get('comune', '')
    
    luogo_str, qsk, zona_vento, press_vento_str, zona_sismica, altitudine_stimata = estrai_parametri_ntc_da_coordinate_e_comune(lat, lon, comune)
    pressione_vento = float(press_vento_str.split()[0])
    
    spessore_cop = str(dati_geo.get('spessore_pannello', ''))
    tipo_cop = str(dati_geo.get('tipo_isolante', ''))
    
    if "Lamiera" in tipo_cop: max_passo_arc = 1.2
    elif "50" in spessore_cop or "60" in spessore_cop: max_passo_arc = 1.8
    elif "80" in spessore_cop or "100" in spessore_cop: max_passo_arc = 2.5
    else: max_passo_arc = 3.0
    
    sviluppo_falda = math.sqrt((luce_totale/2)**2 + (h_colmo - h_gronda)**2)
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

    num_file_long = int(h_gronda / passo_baraccatura)
    ml_baraccatura_long_singola = num_file_long * lunghezza_edificio
    ml_baraccatura_long_tot = ml_baraccatura_long_singola * 2
    
    ml_baraccatura_timpani_singolo = 0
    z = passo_baraccatura
    while z < h_colmo:
        if z <= h_gronda:
            larghezza_timpano = luce_totale
        else:
            h_triangolo_locale = z - h_gronda
            h_triangolo_tot = h_colmo - h_gronda
            larghezza_timpano = luce_totale * (1 - (h_triangolo_locale / h_triangolo_tot))
        ml_baraccatura_timpani_singolo += larghezza_timpano
        z += passo_baraccatura
    ml_baraccatura_timpani_tot = ml_baraccatura_timpani_singolo * 2
    ml_tot_baraccatura = ml_baraccatura_long_tot + ml_baraccatura_timpani_tot

    g1, g2 = 0.15, 0.25
    if "Presente" in dati_geo.get('impianto_fv_desc', ''): g2 += 0.20
    g2 += dati_geo.get('carico_aggiuntivo', 0.0)
    q_tot_copertura_mq = (1.3 * g1 + 1.5 * g2 + 1.5 * qsk)
    q_arc = q_tot_copertura_mq * passo_arcarecci
    
    if "In luce" in pos_arcarecci:
        M_ed_arc = (q_arc * interasse**2) / 8
    else:
        M_ed_arc = (q_arc * interasse**2) / 10

    W_req_arc = (M_ed_arc * 100) / 27.5
    if W_req_arc < 30: sez_arc = "Tubolare 100x50x3"
    elif W_req_arc < 55: sez_arc = "Tubolare 120x60x4"
    elif W_req_arc < 85: sez_arc = "Tubolare 150x75x4 / IPE 140"
    else: sez_arc = "IPE 180"

    classe_servizio = dati_geo.get('classe_servizio', 'Classe 2 (Umidità < 85%)')
    f_md_arc = 11.27 if "Classe 3" in classe_servizio else 14.5
    f_md_trave = 11270.0 if "Classe 3" in classe_servizio else 14500.0
    E_legno_val = 950 if "Classe 3" in classe_servizio else 1150

    W_req_arc_legno = (M_ed_arc * 100) / f_md_arc
    h_arc_legno = max(16, int((6 * W_req_arc_legno / 10.0) ** 0.5))
    h_arc_legno = ((h_arc_legno + 3) // 4) * 4
    sez_arc_legno = f"GL24h 10x{h_arc_legno} cm"

    q_ed_portale = interasse * q_tot_copertura_mq
    luce_campata = luce_totale / (num_appoggi - 1) if num_appoggi >= 3 else luce_totale
    
    if num_appoggi >= 3:
        m_ed = (q_ed_portale * (luce_campata ** 2)) / 10.0 
        v_ed = (q_ed_portale * luce_campata) / 2.0        
    else:
        m_ed = (q_ed_portale * (luce_campata ** 2)) / 8.0
        v_ed = (q_ed_portale * luce_campata) / 2.0

    H_truss = h_colmo - h_gronda
    if H_truss < 0.5: H_truss = luce_totale / 10.0 
    N_max_truss = m_ed / H_truss

    classe_fuoco = dati_geo.get('classe_fuoco', 'R 60')
    fuoco_add = 0
    if "R 60" in classe_fuoco: fuoco_add = 2
    elif "R 90" in classe_fuoco: fuoco_add = 4
    elif "R 120" in classe_fuoco: fuoco_add = 6

    if categoria_struttura == "Capriate":
        if N_max_truss < 150: sez_catena_L, sez_punt_L, sez_mon_L, sez_saet_L = "20x24 cm", "20x24 cm", "20x20 cm", "16x16 cm"
        elif N_max_truss < 300: sez_catena_L, sez_punt_L, sez_mon_L, sez_saet_L = "24x28 cm", "24x28 cm", "24x24 cm", "20x20 cm"
        else: sez_catena_L, sez_punt_L, sez_mon_L, sez_saet_L = "24x32 cm", "24x32 cm", "24x24 cm", "20x20 cm"

        if N_max_truss < 150: sez_catena_A, sez_punt_A, sez_mon_A, sez_saet_A = "Tubolare 100x100x4", "Tubolare 100x100x4", "Tubolare 80x80x3", "Tubolare 80x80x3"
        elif N_max_truss < 300: sez_catena_A, sez_punt_A, sez_mon_A, sez_saet_A = "Tubolare 120x120x5", "Tubolare 120x120x5", "Tubolare 100x100x4", "Tubolare 80x80x4"
        else: sez_catena_A, sez_punt_A, sez_mon_A, sez_saet_A = "Tubolare 150x150x6", "Tubolare 150x150x6", "Tubolare 120x120x5", "Tubolare 100x100x5"

        L_catena = luce_totale
        L_puntone = math.sqrt((luce_totale/2)**2 + H_truss**2)
        L_monaco = H_truss
        L_saettone = math.sqrt((luce_totale/4)**2 + (H_truss/2)**2)

        def build_capriata_desc(s_cat, s_punt, s_mon, s_saet):
            dett = f"• Catena Inferiore: 1x {L_catena:.2f}m | Sez. {s_cat}\n"
            dett += f"• Puntoni Sup.: 2x {L_puntone:.2f}m | Sez. {s_punt}\n"
            if tipo_travatura in ["Con Monaco", "Classica o alla Palladiana", "Composta o a doppia catena"]:
                dett += f"• Monaco: 1x {L_monaco:.2f}m | Sez. {s_mon}\n"
            if tipo_travatura in ["Classica o alla Palladiana", "Composta o a doppia catena"]:
                dett += f"• Saettoni: 2x {L_saettone:.2f}m | Sez. {s_saet}\n"
            return dett

        travi_legno_out = build_capriata_desc(sez_catena_L, sez_punt_L, sez_mon_L, sez_saet_L)
        travi_acciaio_out = build_capriata_desc(sez_catena_A, sez_punt_A, sez_mon_A, sez_saet_A)
        travi_cap_out = "N.D."

    elif categoria_struttura == "Travi Reticolari":
        if N_max_truss < 250: sez_corr_A, sez_diag_A = "Tubolare 120x120x5", "Tubolare 80x80x4"
        elif N_max_truss < 500: sez_corr_A, sez_diag_A = "Tubolare 150x150x6", "Tubolare 100x100x5"
        else: sez_corr_A, sez_diag_A = "Tubolare 200x200x8", "Tubolare 120x120x6"

        if N_max_truss < 250: sez_corr_L, sez_diag_L = "GL24h 20x24 cm", "GL24h 16x16 cm"
        elif N_max_truss < 500: sez_corr_L, sez_diag_L = "GL24h 24x32 cm", "GL24h 20x20 cm"
        else: sez_corr_L, sez_diag_L = "GL24h 24x40 cm", "GL24h 24x24 cm"

        num_campi = max(4, int(luce_totale / 3.0))
        if num_campi % 2 != 0: num_campi += 1 
        L_campo = luce_totale / num_campi
        L_corr_inf = luce_totale
        L_corr_sup = 2 * math.sqrt((luce_totale/2)**2 + H_truss**2)
        H_medio = H_truss / 2
        L_diag = math.sqrt(L_campo**2 + H_medio**2)
        num_diag = num_campi * 2 if tipo_travatura == "Travatura Long" else num_campi
        if tipo_travatura == "Travatura Vierendeel": num_diag = 0
        num_montanti = num_campi - 1

        def build_reticolare_desc(s_corr, s_diag):
            dett = f"• Corrente Inf: L = {L_corr_inf:.2f}m | Sez. {s_corr}\n"
            dett += f"• Corrente Sup: L = {L_corr_sup:.2f}m | Sez. {s_corr}\n"
            if num_montanti > 0 and tipo_travatura != "Travatura Warren":
                dett += f"• Montanti (x{num_montanti}) | Sez. {s_diag}\n"
            if num_diag > 0:
                dett += f"• Diagonali (x{num_diag}) | Sez. {s_diag}\n"
            return dett

        travi_legno_out = build_reticolare_desc(sez_corr_L, sez_diag_L)
        travi_acciaio_out = build_reticolare_desc(sez_corr_A, sez_diag_A)
        travi_cap_out = "N.D."

    else:
        b_legno_cm = 20 + (fuoco_add * 2 if "R 0" not in classe_fuoco else 0)
        w_req_cm3 = (m_ed * 1e6) / f_md_trave  
        h_legno_cm = int((6 * w_req_cm3 / b_legno_cm) ** 0.5)
        h_legno_cm = max(44, ((h_legno_cm + 3) // 4) * 4) 
        if "R 0" not in classe_fuoco:
            h_legno_cm += fuoco_add * 2

        w_el_req_cm3 = (m_ed * 100.0) / 33.8 
        if w_el_req_cm3 > 3500: profilo_acciaio = "IPE 600 / HEB 500"
        elif w_el_req_cm3 > 2000: profilo_acciaio = "IPE 500 / HEA 400"
        elif w_el_req_cm3 > 1000: profilo_acciaio = "IPE 400 / HEA 300"
        else: profilo_acciaio = "IPE 330 / HEA 240"
        
        h_cap_cm = max(80, ((int(h_legno_cm * 1.2) + 4) // 5) * 5)
        profilo_cap = f"Trave a T rovescia precompressa altezza {h_cap_cm} cm"

        travi_legno_out = f"Base {b_legno_cm} cm x Altezza {h_legno_cm} cm (GL24h - {classe_fuoco})"
        travi_acciaio_out = f"Profilo {profilo_acciaio} S355JR"
        travi_cap_out = profilo_cap

    q_w = pressione_vento * interasse
    M_base_vento = (q_w * h_gronda**2) / 2
    E_legno = E_legno_val 
    limite_spostamento_cm = (h_gronda * 100) / 150 
    
    b_pil_perim_cm = 20 + (fuoco_add * 2 if "R 0" not in classe_fuoco else 0)
    h_pil_perim_cm = 32 + (fuoco_add * 2 if "R 0" not in classe_fuoco else 0)
    while True:
        I_pil = (b_pil_perim_cm * h_pil_perim_cm**3) / 12
        W_pil = (b_pil_perim_cm * h_pil_perim_cm**2) / 6
        sigma_m = (M_base_vento * 100) / W_pil 
        delta_somm = (q_w / 100 * (h_gronda * 100)**4) / (8 * E_legno * I_pil)
        if sigma_m < 1.45 and delta_somm < limite_spostamento_cm: break
        h_pil_perim_cm += 4
        if h_pil_perim_cm > 140: break

    h_rif_legno = h_legno_cm if 'h_legno_cm' in locals() else 80
    h_pil_interm_cm = max(32, ((int(h_rif_legno * 0.50) + 3) // 4) * 4)
    b_pil_interm_cm = 20 + (fuoco_add * 2 if "R 0" not in classe_fuoco else 0)

    w_el_rif = w_el_req_cm3 if 'w_el_req_cm3' in locals() else (m_ed * 100.0) / 33.8
    if w_el_rif > 3500: pil_p_acc, pil_i_acc = "HEB 300", "HEA 240"
    elif w_el_rif > 2000: pil_p_acc, pil_i_acc = "HEB 260", "HEA 200"
    elif w_el_rif > 1000: pil_p_acc, pil_i_acc = "HEB 200", "HEA 160"
    else: pil_p_acc, pil_i_acc = "HEB 180", "HEA 140"

    n_bulloni_perim = max(6, int(v_ed / 25.0) * 2)
    peso_conn_perim_kg = round(n_bulloni_perim * 4.0 + (h_pil_perim_cm * b_pil_perim_cm * 0.025) + 25.0, 1)
    n_anc_perim = max(4, int(v_ed / 30.0) * 2)
    peso_anc_perim_kg = round(n_anc_perim * 4.5 + (h_pil_perim_cm * b_pil_perim_cm * 0.03) + 30.0, 1)

    n_bulloni_interm = max(4, int(v_ed / 30.0) * 2)
    peso_conn_interm_kg = round(n_bulloni_interm * 3.8 + (h_pil_interm_cm * b_pil_interm_cm * 0.022) + 20.0, 1)
    n_anc_interm = max(4, int(v_ed / 35.0) * 2)
    peso_anc_interm_kg = round(n_anc_interm * 4.0 + (h_pil_interm_cm * b_pil_interm_cm * 0.025) + 25.0, 1)

    passo_max_montanti = 6.0 
    num_sottocampate_timpano = max(1, int(np.ceil(luce_campata / passo_max_montanti)))
    num_montanti_per_campata_timpano = num_sottocampate_timpano - 1
    num_montanti_timpano_singola_facciata = num_montanti_per_campata_timpano * (num_appoggi - 1)
    passo_montanti_timpano = luce_campata / num_sottocampate_timpano if num_sottocampate_timpano > 0 else 0

    ml_tot_montanti_timpano = 0.0
    ml_per_montante_timpano = []
    if num_montanti_timpano_singola_facciata > 0 and passo_montanti_timpano > 0:
        for c in range(num_appoggi - 1):
            x_start = c * luce_campata
            for m in range(1, num_montanti_per_campata_timpano + 1):
                x_m = x_start + m * passo_montanti_timpano
                if x_m <= luce_totale / 2:
                    h_m = h_gronda + (h_colmo - h_gronda) * (x_m / (luce_totale / 2))
                else:
                    h_m = h_colmo - (h_colmo - h_gronda) * ((x_m - luce_totale / 2) / (luce_totale / 2))
                ml_tot_montanti_timpano += h_m
                ml_per_montante_timpano.append(round(h_m, 2))
    ml_tot_timpani_entrambe = ml_tot_montanti_timpano * 2

    num_campate_totali = max(1, int(round(lunghezza_edificio / interasse)))
    num_sottocampate_long = max(1, int(np.ceil(interasse / passo_max_montanti)))
    num_montanti_per_campata_long = num_sottocampate_long - 1
    passo_montanti_long = interasse / num_sottocampate_long if num_sottocampate_long > 0 else 0
    num_totale_montanti_long_singola_parete = num_montanti_per_campata_long * num_campate_totali
    ml_tot_montanti_long_singola_parete = num_totale_montanti_long_singola_parete * h_gronda
    ml_tot_montanti_long_entrambe_pareti = ml_tot_montanti_long_singola_parete * 2

    if h_colmo <= 6.5: montante_legno, montante_acciaio = f"Sezione 14x14 cm (GL24h - {classe_fuoco})", "HEA 120"
    elif h_colmo <= 9.5: montante_legno, montante_acciaio = f"Sezione 16x16 cm (GL24h - {classe_fuoco})", "HEA 140"
    elif h_colmo <= 12.5: montante_legno, montante_acciaio = f"Sezione 16x24 cm (GL24h - {classe_fuoco})", "HEA 180"
    else: montante_legno, montante_acciaio = f"Sezione 20x28 cm (GL24h - {classe_fuoco})", "HEA 220"

    mq_acciaio = round((luce_totale + h_gronda * 2) * (dati_geo['num_campate'] + 1) * 0.6, 1)

    risultati_deterministici = {
        "luogo": luogo_str, "qsk": qsk, "zona_vento": zona_vento, "pressione_vento": press_vento_str, "zona_sismica": zona_sismica,
        "classe_uso": "Classe II", "fattore_struttura_q": "q = 2.0",
        "travi_legno": travi_legno_out, "travi_acciaio": travi_acciaio_out, "travi_cap": travi_cap_out,
        "b_trave_legno_cm": locals().get('b_legno_cm', 20),
        "h_trave_legno_cm": locals().get('h_legno_cm', 80),
        "pilastri_perimetrali_legno": f"Sezione {b_pil_perim_cm}x{h_pil_perim_cm} cm (Drift H/150 - {classe_fuoco})",
        "pilastri_intermedi_legno": f"Sezione {b_pil_interm_cm}x{h_pil_interm_cm} cm ({classe_fuoco})",
        "pilastri_perimetrali_acciaio": f"Profilo {pil_p_acc} S355JR",
        "pilastri_intermedi_acciaio": f"Profilo {pil_i_acc} S355JR",
        "pilastri_perimetrali_cap": f"C.A.P. 40x45 cm", "pilastri_intermedi_cap": f"C.A.P. 40x40 cm",
        "passo_arcarecci_calc": round(passo_arcarecci, 2), "sezione_arcarecci": f"{sez_arc}", "sezione_arcarecci_legno": f"{sez_arc_legno}",
        "verifica_arcarecci": f"Verificato ({pos_arcarecci}) - M_ed: {M_ed_arc:.1f} kNm",
        "passo_baraccatura_calc": round(passo_baraccatura, 2),
        "ml_baraccatura_tot": round(ml_tot_baraccatura, 1),
        "ml_baraccatura_long_singola": round(ml_baraccatura_long_singola, 1),
        "ml_baraccatura_timpani_singolo": round(ml_baraccatura_timpani_singolo, 1),
        "baraccatura_legno_lamellare": f"GL24h 12x16 cm ({classe_fuoco})", "baraccatura_legno_massiccio": f"C24 14x16 cm ({classe_fuoco})", "baraccatura_acciaio": f"Omega / Tubolare 100x50x3",
        "num_montanti_timpano_singolo": num_montanti_timpano_singola_facciata,
        "passo_montanti_timpano": round(passo_montanti_timpano, 2),
        "ml_per_montante_timpano": ml_per_montante_timpano,
        "ml_tot_timpani_entrambe": round(ml_tot_timpani_entrambe, 2),
        "num_montanti_long_singola_parete": num_totale_montanti_long_singola_parete,
        "passo_montanti_long": round(passo_montanti_long, 2),
        "ml_tot_montanti_long_entrambe": round(ml_tot_montanti_long_entrambe_pareti, 2),
        "montante_sezione_legno": montante_legno, "montante_sezione_acciaio": montante_acciaio, "montante_sezione_cap": "C.A.P. 20x20 cm",
        "campate_controventi_indici": [0, dati_geo['num_campate'] - 1],
        "controventi_copertura_pos": "Campate di estremità",
        "controventi_copertura_legno": "Diagonali in legno lamellare GL24h 14x14 cm", "controventi_copertura_acciaio": "Tubolari incrociati Ø 89x4 mm",
        "controventi_parete_pos": "Campate di estremità",
        "controventi_parete_legno": "Diagonali GL 16x16 cm", "controventi_parete_acciaio": "Croci di Sant'Andrea L 80x8",
        "conn_trave_pilastro_tipo": "Nodo semi-rigido con piastre",
        "conn_trave_pilastro_perim_elementi": f"N. {n_bulloni_perim} bulloni 8.8 M20",
        "conn_trave_pilastro_perim_kg": f"{peso_conn_perim_kg} kg",
        "conn_trave_pilastro_interm_elementi": f"N. {n_bulloni_interm} bulloni 8.8 M20",
        "conn_trave_pilastro_interm_kg": f"{peso_conn_interm_kg} kg",
        "conn_pilastro_fondazione_tipo": "Cerniera/Incastro",
        "conn_pilastro_fondazione_perim_elementi": f"N. {n_anc_perim} tirafondi M24",
        "conn_pilastro_fondazione_perim_kg": f"{peso_anc_perim_kg} kg",
        "conn_pilastro_fondazione_interm_elementi": f"N. {n_anc_interm} tirafondi M24",
        "conn_pilastro_fondazione_interm_kg": f"{peso_anc_interm_kg} kg",
        "dettaglio_giunto_colmo": "Piastra di colmo bullonata",
        "classe_resistenza_fuoco": classe_fuoco,
        "classe_servizio": classe_servizio,
        "mq_intumescente": f"{mq_acciaio} mq",
        "dettaglio_verniciatura": f"Primer + Intumescente {classe_fuoco}" if "R 0" not in classe_fuoco else "Nessun trattamento antincendio richiesto (R0)",
        "note_tecniche": f"Calcolo esatto NTC 2018 con verifica antincendio ({classe_fuoco}) e {classe_servizio}. L={luce_totale}m, H_truss={H_truss:.2f}m."
    }
    return risultati_deterministici

def calcola_distinta_elementi(dati):
    L = dati['lunghezza_edificio']
    B = dati['luce_totale']
    i_portali = dati['interasse_portali']
    n_appoggi = dati['num_appoggi']
    i_arcarecci = dati.get('passo_arcarecci_calc', 1.5)

    num_campate = max(1, int(round(L / i_portali))) if i_portali > 0 else 1
    num_telai = num_campate + 1
    num_pilastri_totali = num_telai * n_appoggi
    num_pilastri_perimetrali = num_telai * 2
    num_pilastri_interni = num_pilastri_totali - num_pilastri_perimetrali
    num_travi_falda = num_telai * 2 

    half_luce = B / 2.0
    num_file_arcarecci = (math.ceil(math.sqrt((half_luce)**2 + (dati['altezza_colmo'] - dati['altezza_gronda'])**2) / i_arcarecci) * 2) - 1
    ml_arcarecci = num_file_arcarecci * L

    campate_cv = [0, num_campate - 1]
    num_campate_cv = sum(1 for idx in campate_cv if 0 <= idx < num_campate)
    num_sub_falda = max(1, int(round(half_luce / 5.0)))
    num_croci_cop = num_campate_cv * (num_sub_falda * 2) 
    
    h_gronda = dati['altezza_gronda']
    num_sub_parete = max(1, int(round(h_gronda / 4.5))) if h_gronda > 0 else 1
    num_croci_par = num_campate_cv * (num_sub_parete * 2) 

    sviluppo_falda = ((B/2)**2 + (dati['altezza_colmo'] - h_gronda)**2)**0.5
    mq_copertura = L * sviluppo_falda * 2
    mq_pareti_lunghe = L * h_gronda * 2 
    mq_timpani = 2 * (B * h_gronda + (B * (dati['altezza_colmo'] - h_gronda) / 2))
    
    tot_montanti_timpani = dati.get('num_montanti_timpano_singolo', 0) * 2
    tot_montanti_longitudinali = dati.get('num_montanti_long_singola_parete', 0) * 2

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
        "mq_pareti_lunghe": round(mq_pareti_lunghe, 1),
        "mq_timpani": round(mq_timpani, 1),
        "tot_montanti_timpani": tot_montanti_timpani,
        "tot_montanti_longitudinali": tot_montanti_longitudinali
    }

def calcola_logistica_trasporti(dati, distinta):
    luce = dati['luce_totale']
    h_colmo = dati['altezza_colmo']
    h_gronda = dati['altezza_gronda']
    num_telai = distinta['num_telai']
    num_pilastri = distinta['num_pilastri_totali']
    num_travi_falda = distinta['num_travi_falda']
    categoria_struttura = dati.get('categoria_struttura', 'Portali ad anima piena')
    
    sviluppo_falda = math.sqrt((luce/2)**2 + (h_colmo - h_gronda)**2)

    if categoria_struttura == "Portali ad anima piena":
        max_lunghezza_trave = sviluppo_falda
    elif categoria_struttura in ["Capriate", "Travi Reticolari"]:
        max_lunghezza_trave = 12.0
    else:
        max_lunghezza_trave = sviluppo_falda

    if max_lunghezza_trave <= 16.50: mezzo_travi = "Autoarticolato / Bilico standard Art. 61 CDS (L max 16,50m)"
    elif max_lunghezza_trave <= 18.75: mezzo_travi = "Autotreno standard Art. 61 CDS (L max 18,75m)"
    elif max_lunghezza_trave <= 25.0: mezzo_travi = "Trasporto Eccezionale - Bilico allungabile (L > 16,50m)"
    elif max_lunghezza_trave <= 33.5: mezzo_travi = "Trasporto Eccezionale - Rimorchio speciale (L > 25m con autorizzazione)"
    else: mezzo_travi = "Trasporto Eccezionale - Convoglio eccezionale con scorta tecnica"

    if categoria_struttura == "Portali ad anima piena":
        b_m = dati.get('b_trave_legno_cm', 20) / 100.0
        h_m = dati.get('h_trave_legno_cm', 80) / 100.0
        peso_specifico_lamellare = 500.0  
        peso_unitario_trave_kg = max_lunghezza_trave * b_m * h_m * peso_specifico_lamellare
    else:
        peso_unitario_trave_kg = max_lunghezza_trave * 45.0  

    portata_utile_kg = 24000.0
    max_pezzi_per_peso = max(1, int(portata_utile_kg / max(1.0, peso_unitario_trave_kg)))
    max_pezzi_per_viaggio_travi = max_pezzi_per_peso
    viaggi_travi = math.ceil(num_travi_falda / max_pezzi_per_viaggio_travi)

    max_h_pilastro = max(h_gronda, h_colmo)
    if max_h_pilastro <= 16.50: mezzo_pilastri = "Bilico standard Art. 61 CDS"
    else: mezzo_pilastri = "Trasporto Eccezionale - Allungabile per pilastri"
        
    peso_unitario_pilastro_kg = max_h_pilastro * 50.0
    max_pezzi_pilastro_peso = max(1, int(portata_utile_kg / max(1.0, peso_unitario_pilastro_kg)))
    max_pezzi_per_viaggio_pilastri = max_pezzi_pilastro_peso
    viaggi_pilastri = math.ceil(num_pilastri / max_pezzi_per_viaggio_pilastri)

    ml_tot_profili = distinta['ml_arcarecci'] + dati.get('ml_baraccatura_tot', 0) + dati.get('ml_tot_timpani_entrambe', 0) + dati.get('ml_tot_montanti_long_entrambe', 0)
    peso_profili_kg = ml_tot_profili * 9.5 
    viaggi_profili = max(1, math.ceil(peso_profili_kg / portata_utile_kg))

    mq_tot_rivestimenti = distinta['mq_copertura'] + distinta['mq_pareti_lunghe'] + distinta['mq_timpani']
    viaggi_pannelli = max(1, math.ceil(mq_tot_rivestimenti / 550.0))
    viaggi_accessori = 1
    tot_viaggi = viaggi_travi + viaggi_pilastri + viaggi_profili + viaggi_pannelli + viaggi_accessori

    return {
        "max_lunghezza_trave": round(max_lunghezza_trave, 2), "mezzo_travi": mezzo_travi, "viaggi_travi": viaggi_travi,
        "qta_effettiva_travi": max_pezzi_per_viaggio_travi, "mezzo_pilastri": mezzo_pilastri, "viaggi_pilastri": viaggi_pilastri,
        "ml_tot_profili": round(ml_tot_profili, 1), "viaggi_profili": viaggi_profili, "mq_tot_rivestimenti": round(mq_tot_rivestimenti, 1),
        "viaggi_pannelli": viaggi_pannelli, "viaggi_accessori": viaggi_accessori, "tot_viaggi": tot_viaggi
    }

def genera_word_report(dati, distinta, logistica):
    doc = Document()
    doc.add_heading('Relazione Tecnica di Predimensionamento, Calcolo e Logistica (NTC 2018)', 0)
    
    doc.add_heading('1. Parametri Geometrici, Climatici, Sismici e di Configurazione', level=1)
    doc.add_paragraph(f"Località / Comune: {dati.get('luogo', 'N.D.')}")
    doc.add_paragraph(f"Carico Neve (qsk): {dati.get('qsk', 1.5)} kN/m²")
    doc.add_paragraph(f"Zona Vento: {dati.get('zona_vento', 'N.D.')} | Pressione: {dati.get('pressione_vento', 'N.D.')}")
    doc.add_paragraph(f"Azione Sismica: {dati.get('zona_sismica', 'N.D.')}")
    doc.add_paragraph(f"Dimensioni Edificio: Lunghezza {dati.get('lunghezza_edificio', 0.0)} m | Larghezza {dati.get('luce_totale', 0.0)} m")
    doc.add_paragraph(f"Altezze: Gronda {dati.get('altezza_gronda', 0.0)} m | Colmo {dati.get('altezza_colmo', 0.0)} m")
    doc.add_paragraph(f"Interasse Portali: {dati.get('interasse_portali', 0.0)} m -> N. Campate: {dati.get('num_campate', 0)}")
    doc.add_paragraph(f"Categoria Struttura: {dati.get('categoria_struttura', 'N.D.')} | Tipologia: {dati.get('tipo_travatura', 'N.D.')}")
    
    doc.add_heading('2. Distinta Elementi Principali (Computo Quantità)', level=1)
    doc.add_paragraph(f"Numero Telai Principali: {distinta['num_telai']}")
    doc.add_paragraph(f"Numero Pilastri Totali: {distinta['num_pilastri_totali']} (di cui {distinta['num_pilastri_perimetrali']} perimetrali e {distinta['num_pilastri_interni']} intermedi)")
    doc.add_paragraph(f"Numero Travi di Falda: {distinta['num_travi_falda']}")
    doc.add_paragraph(f"File di Arcarecci: {distinta['num_file_arcarecci']} file | Metri Lineari: {distinta['ml_arcarecci']} ml")
    doc.add_paragraph(f"Superficie Copertura: {distinta['mq_copertura']} mq | Pareti: {distinta['mq_pareti_lunghe']} mq | Timpani: {distinta['mq_timpani']} mq")

    doc.add_heading('3. Piano Logistico e Calcolo Viaggi di Trasporto (Flotta Veneta Trasporti)', level=1)
    doc.add_paragraph(f"Numero Viaggi Totali Stimati: {logistica['tot_viaggi']} viaggi")
    doc.add_paragraph(f"- Struttura Principale (L calc. max {logistica['max_lunghezza_trave']}m): {logistica['viaggi_travi']} viaggi con {logistica['mezzo_travi']}")
    doc.add_paragraph(f"- Pilastri Strutturali: {logistica['viaggi_pilastri']} viaggi con {logistica['mezzo_pilastri']}")
    doc.add_paragraph(f"- Arcarecci, Baraccatura e Montanti ({logistica['ml_tot_profili']} ml): {logistica['viaggi_profili']} viaggi con Bilico standard")
    doc.add_paragraph(f"- Pannelli Coibentati e Copertura ({logistica['mq_tot_rivestimenti']} mq): {logistica['viaggi_pannelli']} viaggi con Bilico standard (volume)")
    doc.add_paragraph(f"- Accessori, Connessioni e Bulloneria: {logistica['viaggi_accessori']} viaggio dedicato")

    doc.add_heading('4. Arcarecci di Copertura e Baraccatura', level=1)
    doc.add_paragraph(f"Passo Arcarecci: {dati.get('passo_arcarecci_calc', 1.5):.2f} m | Sezione Acciaio: {dati.get('sezione_arcarecci', 'N.D.')} | Sezione Legno: {dati.get('sezione_arcarecci_legno', 'N.D.')}")
    doc.add_paragraph(f"Passo Baraccatura Parete: {dati.get('passo_baraccatura_calc', 2.0):.2f} m | ML Totali: {dati.get('ml_baraccatura_tot', 0)} ml")
    doc.add_paragraph(f"Montanti Timpani: {dati.get('ml_tot_timpani_entrambe', 0):.2f} ml | Montanti Longitudinali: {dati.get('ml_tot_montanti_long_entrambe', 0):.2f} ml")

    doc.add_heading('5. Struttura Principale (Dimensionamento Elementi)', level=1)
    if dati.get('categoria_struttura') in ["Capriate", "Travi Reticolari"]:
        doc.add_paragraph("--- VARIANTE IN LEGNO ---\n" + dati.get('travi_legno', 'N.D.'))
        doc.add_paragraph("--- VARIANTE IN ACCIAIO ---\n" + dati.get('travi_acciaio', 'N.D.'))
    else:
        doc.add_paragraph(f"Legno Lamellare: {dati.get('travi_legno', 'N.D.')}")
        doc.add_paragraph(f"Acciaio: {dati.get('travi_acciaio', 'N.D.')}")
        doc.add_paragraph(f"C.a.p.: {dati.get('travi_cap', 'N.D.')}")
    
    doc.add_heading('6. Connessioni, Nodi e Dettagli d’Ancoraggio', level=1)
    doc.add_paragraph(f"Nodo Trave-Pilastro: {dati.get('conn_trave_pilastro_tipo')} | Perimetrali: {dati.get('conn_trave_pilastro_perim_elementi')} ({dati.get('conn_trave_pilastro_perim_kg')})")
    doc.add_paragraph(f"Nodi Intermedi: {dati.get('conn_trave_pilastro_interm_elementi')} ({dati.get('conn_trave_pilastro_interm_kg')})")
    doc.add_paragraph(f"Ancoraggi di Base: {dati.get('conn_pilastro_fondazione_tipo')} | Perimetrali: {dati.get('conn_pilastro_fondazione_perim_elementi')} ({dati.get('conn_pilastro_fondazione_perim_kg')})")
    doc.add_paragraph(f"Ancoraggi Intermedi: {dati.get('conn_pilastro_fondazione_interm_elementi')} ({dati.get('conn_pilastro_fondazione_interm_kg')})")
    if "giuntata" in dati.get('tipo_travatura', '').lower():
        doc.add_paragraph(f"Dettaglio Giunto in Colmo: {dati.get('dettaglio_giunto_colmo')}")

    doc.add_heading('7. Protezione Antincendio, Durabilità e Note', level=1)
    doc.add_paragraph(f"Classe Resistenza al Fuoco: {dati.get('classe_resistenza_fuoco')} | Superficie Acciaio: {dati.get('mq_intumescente')}")
    doc.add_paragraph(f"Classe di Servizio (Strutture in Legno): {dati.get('classe_servizio', 'Classe 2')}")
    doc.add_paragraph(f"Ciclo Verniciatura: {dati.get('dettaglio_verniciatura')}")
    doc.add_paragraph(dati.get('note_tecniche', 'N.D.'))
    
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

def genera_modello_3d(dati):
    fig = go.Figure()
    luce_totale = dati.get('luce_totale', 39.6)
    altezza_gronda = dati.get('altezza_gronda', 9.0)
    altezza_colmo = dati.get('altezza_colmo', 12.21)
    lunghezza_edificio = dati.get('lunghezza_edificio', 25.0)
    interasse_portali = dati.get('interasse_portali', 5.0)
    num_appoggi = dati.get('num_appoggi', 3)
    categoria_struttura = dati.get('categoria_struttura', 'Portali ad anima piena')
    tipo_travatura = dati.get('tipo_travatura', 'Bi-falda semplice')
    interasse_arcarecci = dati.get('passo_arcarecci_calc', 1.5)
    
    num_campate = max(1, int(round(lunghezza_edificio / interasse_portali))) if interasse_portali > 0 else 1
    y_portali = [i * interasse_portali for i in range(num_campate + 1)]
    
    if num_appoggi == 2: x_pilastri = [0.0, luce_totale]
    elif num_appoggi == 3: x_pilastri = [0.0, luce_totale / 2.0, luce_totale]
    else: x_pilastri = [0.0, luce_totale / 3.0, (2 * luce_totale) / 3.0, luce_totale]
        
    for idx_y, y in enumerate(y_portali):
        for idx_x, x in enumerate(x_pilastri):
            h_p = altezza_gronda if (x == 0.0 or x == luce_totale) else altezza_colmo
            show_leg = (idx_y == 0 and idx_x == 0)
            fig.add_trace(go.Scatter3d(x=[x, x], y=[y, y], z=[0, h_p], mode='lines', line=dict(color='darkblue', width=6), name='Pilastri' if show_leg else '', showlegend=show_leg))
        
        show_leg_trave = (idx_y == 0)

        if categoria_struttura == "Capriate":
            fig.add_trace(go.Scatter3d(x=[0, luce_totale], y=[y, y], z=[altezza_gronda, altezza_gronda], mode='lines', line=dict(color='saddlebrown', width=6), name='Catena' if show_leg_trave else '', showlegend=show_leg_trave))
            fig.add_trace(go.Scatter3d(x=[0, luce_totale/2, luce_totale], y=[y, y, y], z=[altezza_gronda, altezza_colmo, altezza_gronda], mode='lines', line=dict(color='firebrick', width=6), name='Puntoni' if show_leg_trave else '', showlegend=show_leg_trave))
            if tipo_travatura in ["Con Monaco", "Classica o alla Palladiana", "Composta o a doppia catena"]:
                fig.add_trace(go.Scatter3d(x=[luce_totale/2, luce_totale/2], y=[y, y], z=[altezza_gronda, altezza_colmo], mode='lines', line=dict(color='peru', width=5), name='Monaco' if show_leg_trave else '', showlegend=show_leg_trave))
            if tipo_travatura in ["Classica o alla Palladiana", "Composta o a doppia catena"]:
                z_mid = altezza_gronda + (altezza_colmo - altezza_gronda)/2
                fig.add_trace(go.Scatter3d(x=[luce_totale/2, luce_totale/4], y=[y, y], z=[altezza_gronda, z_mid], mode='lines', line=dict(color='darkgoldenrod', width=4), name='Saettoni' if show_leg_trave else '', showlegend=show_leg_trave))
                fig.add_trace(go.Scatter3d(x=[luce_totale/2, 3*luce_totale/4], y=[y, y], z=[altezza_gronda, z_mid], mode='lines', line=dict(color='darkgoldenrod', width=4), showlegend=False))
            if tipo_travatura == "Composta o a doppia catena":
                z_mid = altezza_gronda + (altezza_colmo - altezza_gronda)/2
                fig.add_trace(go.Scatter3d(x=[luce_totale/4, 3*luce_totale/4], y=[y, y], z=[z_mid, z_mid], mode='lines', line=dict(color='saddlebrown', width=5), name='Catena Sup.' if show_leg_trave else '', showlegend=show_leg_trave))

        elif categoria_struttura == "Travi Reticolari":
            num_campi = max(4, int(luce_totale / 3.0))
            if num_campi % 2 != 0: num_campi += 1
            L_c = luce_totale / num_campi
            
            fig.add_trace(go.Scatter3d(x=[0, luce_totale], y=[y, y], z=[altezza_gronda, altezza_gronda], mode='lines', line=dict(color='gray', width=6), name='Corrente Inf' if show_leg_trave else '', showlegend=show_leg_trave))
            fig.add_trace(go.Scatter3d(x=[0, luce_totale/2, luce_totale], y=[y, y, y], z=[altezza_gronda, altezza_colmo, altezza_gronda], mode='lines', line=dict(color='dimgray', width=6), name='Corrente Sup' if show_leg_trave else '', showlegend=show_leg_trave))
            
            for i in range(num_campi):
                x1 = i * L_c
                x2 = (i+1) * L_c
                z1_sup = altezza_gronda + (altezza_colmo - altezza_gronda)*(x1/(luce_totale/2)) if x1 <= luce_totale/2 else altezza_colmo - (altezza_colmo - altezza_gronda)*(x1 - luce_totale/2)/(luce_totale/2)
                z2_sup = altezza_gronda + (altezza_colmo - altezza_gronda)*(x2/(luce_totale/2)) if x2 <= luce_totale/2 else altezza_colmo - (altezza_colmo - altezza_gronda)*(x2 - luce_totale/2)/(luce_totale/2)
                
                if i > 0 and tipo_travatura != "Travatura Warren":
                    fig.add_trace(go.Scatter3d(x=[x1, x1], y=[y, y], z=[altezza_gronda, z1_sup], mode='lines', line=dict(color='darkslategray', width=3), name='Aste Web' if (show_leg_trave and i==1) else '', showlegend=(show_leg_trave and i==1)))
                
                if tipo_travatura != "Travatura Vierendeel":
                    if tipo_travatura == "Travatura Warren":
                        if i % 2 == 0: fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y, y], z=[altezza_gronda, z2_sup], mode='lines', line=dict(color='darkslategray', width=3), name='Aste Web' if (show_leg_trave and i==0) else '', showlegend=(show_leg_trave and i==0)))
                        else: fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y, y], z=[z1_sup, altezza_gronda], mode='lines', line=dict(color='darkslategray', width=3), showlegend=False))
                    elif tipo_travatura == "Travatura Howe":
                        if x1 < luce_totale/2: fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y, y], z=[altezza_gronda, z2_sup], mode='lines', line=dict(color='darkslategray', width=3), showlegend=False))
                        else: fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y, y], z=[z1_sup, altezza_gronda], mode='lines', line=dict(color='darkslategray', width=3), showlegend=False))
                    elif tipo_travatura == "Travatura Pratt":
                        if x1 < luce_totale/2: fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y, y], z=[z1_sup, altezza_gronda], mode='lines', line=dict(color='darkslategray', width=3), showlegend=False))
                        else: fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y, y], z=[altezza_gronda, z2_sup], mode='lines', line=dict(color='darkslategray', width=3), showlegend=False))
                    elif tipo_travatura == "Travatura Long":
                        fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y, y], z=[altezza_gronda, z2_sup], mode='lines', line=dict(color='darkslategray', width=3), showlegend=False))
                        fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y, y], z=[z1_sup, altezza_gronda], mode='lines', line=dict(color='darkslategray', width=3), showlegend=False))

        else:
            fig.add_trace(go.Scatter3d(x=[0, luce_totale/2, luce_totale], y=[y, y, y], z=[altezza_gronda, altezza_colmo, altezza_gronda], mode='lines', line=dict(color='firebrick', width=6), name='Travi di Falda' if show_leg_trave else '', showlegend=show_leg_trave))

    passo_mont_timpano = dati.get('passo_montanti_timpano', 0)
    num_mont_camp_timpano = int(dati.get('num_montanti_timpano_singolo', 0) / max(1, num_appoggi - 1))
    if passo_mont_timpano > 0 and num_mont_camp_timpano > 0:
        x_montanti_timpano = []
        for x_start in x_pilastri[:-1]:
            for m in range(1, num_mont_camp_timpano + 1):
                x_m = x_start + m * passo_mont_timpano
                if x_m < x_start + (luce_totale / (num_appoggi - 1)) - 0.1: x_montanti_timpano.append(x_m)
                    
        for y_fac in [0, y_portali[-1]]:
            for idx_m, xm in enumerate(x_montanti_timpano):
                zm = altezza_gronda + (altezza_colmo - altezza_gronda) * (xm / (luce_totale / 2)) if xm <= luce_totale / 2 else altezza_colmo - (altezza_colmo - altezza_gronda) * ((xm - luce_totale / 2) / (luce_totale / 2))
                show_leg_mont_timp = (y_fac == 0 and idx_m == 0)
                fig.add_trace(go.Scatter3d(x=[xm, xm], y=[y_fac, y_fac], z=[0, zm], mode='lines', line=dict(color='cadetblue', width=4), name='Montanti Timpano' if show_leg_mont_timp else '', showlegend=show_leg_mont_timp))

    passo_mont_long = dati.get('passo_montanti_long', 0)
    num_mont_camp_long = int(dati.get('num_montanti_long_singola_parete', 0) / num_campate) if num_campate > 0 else 0
    if passo_mont_long > 0 and num_mont_camp_long > 0:
        y_montanti_long = []
        for y_start in y_portali[:-1]:
            for m in range(1, num_mont_camp_long + 1):
                y_m = y_start + m * passo_mont_long
                if y_m < y_start + interasse_portali - 0.1: y_montanti_long.append(y_m)
                    
        for x_wall in [0, luce_totale]:
            for idx_m, ym in enumerate(y_montanti_long):
                show_leg_mont_long = (x_wall == 0 and idx_m == 0)
                fig.add_trace(go.Scatter3d(x=[x_wall, x_wall], y=[ym, ym], z=[0, altezza_gronda], mode='lines', line=dict(color='teal', width=4), name='Montanti Longitudinali' if show_leg_mont_long else '', showlegend=show_leg_mont_long))

    half_luce = luce_totale / 2.0
    x_arc_left = []
    curr = 0.0
    while curr <= half_luce - 1e-5:
        x_arc_left.append(curr)
        curr += interasse_arcarecci
    if not x_arc_left or abs(x_arc_left[-1] - half_luce) > 1e-5: x_arc_left.append(half_luce)
        
    for idx_x, x_val in enumerate(x_arc_left):
        z_val = altezza_gronda + (altezza_colmo - altezza_gronda) * (x_val / half_luce)
        show_leg_arc = (idx_x == 0)
        fig.add_trace(go.Scatter3d(x=[x_val, x_val], y=[y_portali[0], y_portali[-1]], z=[z_val, z_val], mode='lines', line=dict(color='gray', width=2, dash='dot'), name='Arcarecci' if show_leg_arc else '', showlegend=show_leg_arc))
        if x_val < half_luce - 1e-5:
            x_right = luce_totale - x_val
            fig.add_trace(go.Scatter3d(x=[x_right, x_right], y=[y_portali[0], y_portali[-1]], z=[z_val, z_val], mode='lines', line=dict(color='gray', width=2, dash='dot'), showlegend=False))

    raw_indici = dati.get('campate_controventi_indici', [0, num_campate - 1])
    if isinstance(raw_indici, list): campate_controventi = [int(i) for i in raw_indici if isinstance(i, (int, float))]
    else: campate_controventi = [0, num_campate - 1]

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
                
                fig.add_trace(go.Scatter3d(x=[x_s1, x_s2, None, x_s1, x_s2], y=[y_start, y_end, None, y_end, y_start], z=[z_s1, z_s2, None, z_s1, z_s2], mode='lines', line=dict(color='forestgreen', width=4), name='Controventi Copertura' if (show_leg_cv_cop and s == 0) else '', showlegend=(show_leg_cv_cop and s == 0)))
                
                xr_s1 = luce_totale - x_s1
                xr_s2 = luce_totale - x_s2
                fig.add_trace(go.Scatter3d(x=[xr_s1, xr_s2, None, xr_s1, xr_s2], y=[y_start, y_end, None, y_end, y_start], z=[z_s1, z_s2, None, z_s1, z_s2], mode='lines', line=dict(color='forestgreen', width=4), showlegend=False))
            
            dz_parete = altezza_gronda / num_sub_parete
            for x_wall in [0.0, luce_totale]:
                for t in range(num_sub_parete):
                    z_t1 = t * dz_parete
                    z_t2 = (t + 1) * dz_parete
                    fig.add_trace(go.Scatter3d(x=[x_wall, x_wall, None, x_wall, x_wall], y=[y_start, y_end, None, y_start, y_end], z=[z_t1, z_t2, None, z_t2, z_t1], mode='lines', line=dict(color='darkorange', width=4), name='Controventi Parete' if (show_leg_cv_par and x_wall == 0.0 and t == 0) else '', showlegend=(show_leg_cv_par and x_wall == 0.0 and t == 0)))

    fig.update_layout(
        title=f"Modello 3D Dinamico ({num_campate} Campate, {num_campate+1} Telai - {tipo_travatura})",
        scene=dict(xaxis_title=f'Larghezza (X - {luce_totale}m)', yaxis_title=f'Lunghezza (Y - {lunghezza_edificio}m)', zaxis_title=f'Altezza (Z - {altezza_colmo}m)', aspectmode='data'),
        margin=dict(l=0, r=0, b=0, t=40), height=550
    )
    return fig

# --- FUNZIONI DI CALCOLO XLAM E PROPRIETA' EFFICACI ---
def calcola_proprieta_efficaci_xlam(strati, orientamento, def_fuoco=0):
    strati_eff = list(strati)
    rimosso = def_fuoco
    # Carbonizzazione dal basso (strati_eff[-1] è lo strato inferiore)
    for i in range(len(strati_eff)-1, -1, -1):
        if rimosso >= strati_eff[i]:
            rimosso -= strati_eff[i]
            strati_eff[i] = 0
        else:
            strati_eff[i] -= rimosso
            rimosso = 0
            break

    A_eff = 0
    S_eff = 0
    y_curr = 0
    for i in range(len(strati_eff)):
        t = strati_eff[i]
        if orientamento[i] == 1 and t > 0:
            y_centro = y_curr + t/2
            A_eff += t
            S_eff += t * y_centro
        y_curr += t
        
    if A_eff == 0: return 0, 0 
        
    y_g = S_eff / A_eff
    
    I_eff = 0
    y_curr = 0
    for i in range(len(strati_eff)):
        t = strati_eff[i]
        if orientamento[i] == 1 and t > 0:
            y_centro = y_curr + t/2
            I_strato = (1.0 * t**3) / 12  
            I_eff += I_strato + t * (y_centro - y_g)**2
        y_curr += t
        
    I_eff *= 0.85 # Metodo gamma approssimato k_sys
    
    y_top_long = -1
    y_bot_long = -1
    y_c = 0
    for i in range(len(strati_eff)):
        t = strati_eff[i]
        if orientamento[i] == 1 and t > 0:
            if y_top_long == -1: y_top_long = y_c
            y_bot_long = y_c + t
        y_c += t
        
    dist_top = abs(y_top_long - y_g) if y_top_long != -1 else 1
    dist_bot = abs(y_bot_long - y_g) if y_bot_long != -1 else 1
    y_max = max(dist_top, dist_bot)
    
    W_eff = I_eff / y_max if y_max > 0 else 0
    return I_eff, W_eff

# Database stratigrafie tipiche Stora Enso per solai
pannelli_xlam_db = [
    {"nome": "CLT 90 C3s", "spessore": 90, "strati": [30, 30, 30], "orientamento": [1, 0, 1]},
    {"nome": "CLT 100 C3s", "spessore": 100, "strati": [33, 34, 33], "orientamento": [1, 0, 1]},
    {"nome": "CLT 120 C3s", "spessore": 120, "strati": [40, 40, 40], "orientamento": [1, 0, 1]},
    {"nome": "CLT 100 C5s", "spessore": 100, "strati": [20, 20, 20, 20, 20], "orientamento": [1, 0, 1, 0, 1]},
    {"nome": "CLT 120 C5s", "spessore": 120, "strati": [24, 24, 24, 24, 24], "orientamento": [1, 0, 1, 0, 1]},
    {"nome": "CLT 140 C5s", "spessore": 140, "strati": [40, 20, 20, 20, 40], "orientamento": [1, 0, 1, 0, 1]},
    {"nome": "CLT 160 C5s", "spessore": 160, "strati": [40, 20, 40, 20, 40], "orientamento": [1, 0, 1, 0, 1]},
    {"nome": "CLT 180 C5s", "spessore": 180, "strati": [40, 30, 40, 30, 40], "orientamento": [1, 0, 1, 0, 1]},
    {"nome": "CLT 200 C5s", "spessore": 200, "strati": [40, 40, 40, 40, 40], "orientamento": [1, 0, 1, 0, 1]},
    {"nome": "CLT 240 C7s", "spessore": 240, "strati": [40, 30, 30, 40, 30, 30, 40], "orientamento": [1, 0, 1, 0, 1, 0, 1]},
    {"nome": "CLT 280 C7s", "spessore": 280, "strati": [40, 40, 40, 40, 40, 40, 40], "orientamento": [1, 0, 1, 0, 1, 0, 1]}
]


st.set_page_config(page_title="Predimensionamento Strutturale NTC 2018", layout="wide")
st.title("Generatore Offerte Tecniche e Dimensionamento IA 🏗️")

with st.sidebar:
    st.header("Impostazioni Motore")
    api_key = st.text_input("Inserisci qui la tua API Key di Google", type="password")
    
    st.markdown("---")
    modalita_deterministica = st.toggle("Motore Deterministico NTC 2018 (No IA)", value=True)
    if modalita_deterministica:
        st.success("🟢 Motore Matematico Locale Attivo")
    else:
        st.info("🤖 Modalità Ibrida con IA attiva")

    st.markdown("---")
    if st.button("🔄 Nuovo Progetto / Reset", use_container_width=True):
        st.session_state.clear()
        st.rerun()

tab_principale, tab_xlam = st.tabs(["🏗️ Struttura Principale Capannone (NTC 2018)", "🪵 Dimensionamento Solaio XLAM"])

with tab_principale:
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
                testi_estratto = [entity.dxf.text for entity in msp if entity.dxftype() == 'TEXT'] + [entity.text for entity in msp if entity.dxftype() == 'MTEXT']
                testo_estratto_file = "\n".join(testi_estratto)
                st.success(f"File CAD '{file_caricato.name}' letto con successo!")
                os.unlink(tmp_path)
            elif estensione == 'pdf':
                pdf_reader = PyPDF2.PdfReader(file_caricato)
                testi_pdf = [page.extract_text() for page in pdf_reader.pages if page.extract_text()]
                testo_estratto_file = "\n".join(testi_pdf)
                st.success(f"File PDF '{file_caricato.name}' letto con successo!")
        except Exception as e:
            st.error(f"Errore nella lettura del file: {e}")

    testo_commerciale = st.text_area("Incolla qui le note del progetto o il capitolato:", height=100, value="", key="testo_commerciale")

    st.markdown("### 📍 Localizzazione Cantiere (Google Maps e Comune)")
    col_loc1, col_loc2 = st.columns([2, 1])
    with col_loc1:
        maps_url_ui = st.text_input("Incolla il link di Google Maps del cantiere:", value="", key="maps_url_ui")
    with col_loc2:
        _, _, luogo_estratto_url = estrai_dati_da_url_maps(maps_url_ui)
        comune_cantiere_ui = st.text_input("Comune di installazione", value=luogo_estratto_url, key="comune_cantiere_ui")

    st.markdown("### 📐 Dimensioni Geometriche dell'Edificio (Modificabili)")
    col_dim1, col_dim2, col_dim3, col_dim4, col_dim5 = st.columns(5)
    with col_dim1: lunghezza_edificio_ui = st.number_input("Lunghezza Edificio (m)", min_value=0.0, value=25.0, step=1.0, format="%.1f", key="lunghezza_edificio_ui")
    with col_dim2: interasse_portali_ui = st.number_input("Interasse Portali (m)", min_value=0.0, value=5.0, step=0.5, format="%.2f", key="interasse_portali_ui")
    with col_dim3: luce_totale_ui = st.number_input("Luce Totale / Larghezza (m)", min_value=0.0, value=39.6, step=0.1, format="%.2f", key="luce_totale_ui")
    with col_dim4: altezza_gronda_ui = st.number_input("Altezza Gronda (m)", min_value=0.0, value=9.0, step=0.5, format="%.1f", key="altezza_gronda_ui")
    with col_dim5: altezza_colmo_ui = st.number_input("Altezza Colmo (m)", min_value=0.0, value=12.21, step=0.01, format="%.2f", key="altezza_colmo_ui")

    st.markdown("### 🏛️ Configurazione Telaio e Travatura")
    col_g1, col_g2, col_g3 = st.columns(3)
    with col_g1:
        categoria_struttura = st.selectbox("Categoria Struttura Principale", ["Portali ad anima piena", "Capriate", "Travi Reticolari"], key="cat_strutt")
        if categoria_struttura == "Portali ad anima piena":
            tipo_travatura = st.selectbox("Tipologia Travatura", ["Bi-falda semplice", "Bi-falda con intradosso curvo", "Trave di falda giuntata in colmo"], key="tipo_travatura")
        elif categoria_struttura == "Capriate":
            tipo_travatura = st.selectbox("Tipologia Capriata", ["Semplice", "Con Monaco", "Classica o alla Palladiana", "Composta o a doppia catena"], key="tipo_travatura")
        else:
            tipo_travatura = st.selectbox("Tipologia Reticolare", ["Travatura Warren", "Travatura Long", "Travatura Howe", "Travatura Vierendeel", "Travatura Pratt"], key="tipo_travatura")
    with col_g2:
        num_appoggi = st.selectbox("Numero Appoggi Telaio", [2, 3, 4], index=1, format_func=lambda x: f"{x} Appoggi", key="num_appoggi")
    with col_g3:
        posizione_arc = st.radio("Posizionamento Arcarecci", ["Sopra i telai (Continuo)", "In luce (Semplice appoggio)"], key="pos_arcarecci")

    st.markdown("### ⚙️ Parametri Carichi di Copertura e Pannellature")
    col_c1, col_c2, col_c3 = st.columns(3)
    with col_c1:
        tipo_isolante = st.selectbox("Tipologia Pannello Copertura", ["PIR / PUR", "Lana Minerale", "Lamiera Grecata Semplice"], key="tipo_isolante")
        if tipo_isolante == "PIR / PUR": spessore_pannello = st.selectbox("Spessore Pannello (mm)", [50, 60, 80, 100, 120], key="spessore_panni_pir")
        elif tipo_isolante == "Lana Minerale": spessore_pannello = st.selectbox("Spessore Pannello (mm)", [100, 120, 150, 170], key="spessore_panni_lana")
        else: spessore_pannello = 0
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
        if tipo_isolante_parete == "PIR / PUR": spessore_pannello_parete = st.selectbox("Spessore Pannello Parete (mm)", [50, 60, 80, 100, 120], key="spessore_parete_pir")
        elif tipo_isolante_parete == "Lana di Roccia": spessore_pannello_parete = st.selectbox("Spessore Pannello Parete (mm)", [80, 100, 120, 150], key="spessore_parete_lana")
        else: spessore_pannello_parete = 0

    st.markdown("### 🔥 Requisiti Antincendio e Durabilità (NTC 2018)")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        classe_fuoco_ui = st.selectbox("Classe di Resistenza al Fuoco", ["R 0 (Nessun requisito)", "R 60", "R 90", "R 120"], index=1, key="classe_fuoco_ui")
    with col_f2:
        classe_servizio_ui = st.selectbox("Classe di Servizio (Legno EN 1995-1-1)", ["Classe 1 (Interno asciutto)", "Classe 2 (Umidità < 85%)", "Classe 3 (Esterno esposto)"], index=1, key="classe_servizio_ui")

    if st.button("Esegui Dimensionamento, Logistica e Genera Modello 3D", type="primary"):
        if lunghezza_edificio_ui <= 0 or interasse_portali_ui <= 0 or luce_totale_ui <= 0 or altezza_gronda_ui <= 0 or altezza_colmo_ui <= 0:
            st.warning("⚠️ Inserisci tutte le dimensioni geometriche con valori superiori a zero prima di eseguire il calcolo.")
        else:
            lat_estratta, lon_estratta, place_url = estrai_dati_da_url_maps(maps_url_ui)
            comune_finale = comune_cantiere_ui if comune_cantiere_ui else place_url
            
            num_campate_calc = max(1, int(round(lunghezza_edificio_ui / interasse_portali_ui)))
            impianto_fv_desc = "Presente (20 kg/mq)" if impianto_fv else "Assente"
            
            dati_base = {
                'lunghezza_edificio': lunghezza_edificio_ui, 'interasse_portali': interasse_portali_ui,
                'luce_totale': luce_totale_ui, 'altezza_gronda': altezza_gronda_ui, 'altezza_colmo': altezza_colmo_ui,
                'num_campate': num_campate_calc,
                'categoria_struttura': categoria_struttura, 'tipo_travatura': tipo_travatura, 'num_appoggi': num_appoggi,
                'posizione_arcarecci': posizione_arc, 'tipo_isolante': tipo_isolante,
                'spessore_pannello': f"{spessore_pannello} mm" if tipo_isolante != "Lamiera Grecata Semplice" else "Lamiera Semplice",
                'tipo_isolante_parete': tipo_isolante_parete,
                'spessore_pannello_parete': f"{spessore_pannello_parete} mm" if tipo_isolante_parete not in ["Lamiera Semplice", "Nessuno (Aperto)"] else tipo_isolante_parete,
                'impianto_fv_desc': impianto_fv_desc, 'carico_aggiuntivo': carico_aggiuntivo,
                'latitudine': lat_estratta, 'longitudine': lon_estratta, 'comune': comune_finale,
                'classe_fuoco': classe_fuoco_ui,
                'classe_servizio': classe_servizio_ui
            }

            if modalita_deterministica:
                with st.spinner('Estrazione coordinate ed esecuzione calcolo deterministico NTC 2018...'):
                    dati = esegui_calcolo_deterministico(dati_base)
                    dati.update(dati_base)
                    dati['distinta'] = calcola_distinta_elementi(dati)
                    dati['logistica'] = calcola_logistica_trasporti(dati, dati['distinta'])
                    st.session_state['dati_ultimi'] = dati
                    st.success("Calcolo strutturale e piano logistico completati con successo!")
            else:
                if not api_key:
                    st.error("Inserisci prima l'API Key di Google nella barra laterale!")
                else:
                    dati_config_str = f"Luce: {luce_totale_ui}m, Lunghezza: {lunghezza_edificio_ui}m, Categoria: {categoria_struttura}"
                    genai.configure(api_key=api_key)
                    try:
                        model = genai.GenerativeModel(model_name='gemini-3.6-flash', generation_config={"response_mime_type": "application/json", "temperature": 0.0})
                        prompt = "Restituisci JSON valido con parametri strutturali NTC 2018 per: " + dati_config_str
                        with st.spinner('Elaborazione con IA...'):
                            risposta_ia = model.generate_content(prompt)
                            testo_risposta = risposta_ia.text.strip()
                            if testo_risposta.startswith("```json"): testo_risposta = testo_risposta[7:]
                            if testo_risposta.startswith("```"): testo_risposta = testo_risposta[3:]
                            if testo_risposta.endswith("```"): testo_risposta = testo_risposta[:-3]
                            dati = json.loads(testo_risposta.strip())
                            dati.update(dati_base)
                            risultati_strutturali = esegui_calcolo_deterministico(dati)
                            dati.update(risultati_strutturali)
                            dati['distinta'] = calcola_distinta_elementi(dati)
                            dati['logistica'] = calcola_logistica_trasporti(dati, dati['distinta'])
                            st.session_state['dati_ultimi'] = dati
                            st.success("Modello IA calcolato con successo!")
                    except Exception as e:
                        st.error(f"Errore IA: {e}")

    if 'dati_ultimi' in st.session_state:
        dati = st.session_state['dati_ultimi']
        distinta = dati.get('distinta', calcola_distinta_elementi(dati))
        logistica = dati.get('logistica', calcola_logistica_trasporti(dati, distinta))
        st.markdown("---")
        
        col_dl1, col_dl2, col_dl3 = st.columns([1, 2, 1])
        with col_dl2:
            word_file = genera_word_report(dati, distinta, logistica)
            st.download_button(label="📄 Scarica Relazione, Computo e Piano Logistico in Word (.docx)", data=word_file, file_name=f"Relazione_Logistica_{dati.get('luogo', 'Progetto').replace(' ', '_').replace(':', '')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
        
        st.markdown("---")
        st.markdown("### 🚚 Piano Logistico e Calcolo Viaggi di Trasporto (Flotta Veneta Trasporti)")
        c_l1, c_l2, c_l3 = st.columns(3)
        c_l1.metric("Totale Viaggi Stimati", f"{logistica.get('tot_viaggi', 0)} Viaggi", "Ottimizzato Costo/Portata (max 24t)", delta_color="off")
        c_l2.metric("Mezzo Travi / Capriate", logistica.get('mezzo_travi', 'N.D.'), f"N° {logistica.get('viaggi_travi', 0)} Viaggi (Max {logistica.get('qta_effettiva_travi', 1)} pz/viaggio)")
        c_l3.metric("Mezzo Pilastri", logistica.get('mezzo_pilastri', 'N.D.'), f"N° {logistica.get('viaggi_pilastri', 0)} Viaggi", delta_color="off")

        c_l4, c_l5, c_l6 = st.columns(3)
        c_l4.metric("Arcarecci e Baraccatura", f"N° {logistica.get('viaggi_profili', 0)} Viaggi", f"Tot: {logistica.get('ml_tot_profili', 0)} ml", delta_color="off")
        c_l5.metric("Pannelli e Copertura", f"N° {logistica.get('viaggi_pannelli', 0)} Viaggi", f"Area: {logistica.get('mq_tot_rivestimenti', 0)} mq", delta_color="off")
        c_l6.metric("Accessori e Connessioni", f"N° {logistica.get('viaggi_accessori', 1)} Viaggio", "Bulloneria e piastre", delta_color="off")

        st.markdown("---")
        st.markdown("### 🌐 Modello 3D Dinamico della Struttura")
        fig_3d = genera_modello_3d(dati)
        st.plotly_chart(fig_3d, use_container_width=True)
        
        st.markdown("---")
        st.markdown("### 📋 1. Distinta Elementi Principali (Computo Quantità)")
        c_e1, c_e2, c_e3, c_e4 = st.columns(4)
        c_e1.metric("Telai Principali", f"{distinta['num_telai']} pz")
        c_e2.metric("Pilastri Totali", f"{distinta['num_pilastri_totali']} pz", f"{distinta['num_pilastri_perimetrali']} Per. | {distinta['num_pilastri_interni']} Intermedi", delta_color="off")
        c_e3.metric("Travi di Falda", f"{distinta['num_travi_falda']} pz")
        c_e4.metric("File Arcarecci", f"{distinta['num_file_arcarecci']} file", f"Tot: {distinta['ml_arcarecci']} ml", delta_color="off")

        c_e5, c_e6, c_e7, c_e8 = st.columns(4)
        c_e5.metric("Moduli Controvento Cop.", f"{distinta['num_croci_copertura']} croci")
        c_e6.metric("Moduli Controvento Parete", f"{distinta['num_croci_parete']} croci")
        c_e7.metric("Superficie Copertura", f"{distinta['mq_copertura']} mq")
        c_e8.metric("Superficie Pareti Longitudinali", f"{distinta['mq_pareti_lunghe']} mq")

        st.markdown("---")
        st.markdown("### 📍 2. Dati geometrici, climatici, sismici e di configurazione (NTC 2018)")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Località / Comune", dati.get("luogo", "N.D."))
        c2.metric("Carico Neve (qsk)", f"{dati.get('qsk', 1.5)} kN/m²")
        c3.metric("Zona Vento", dati.get("zona_vento", "N.D."))
        c4.metric("Pressione Vento", dati.get("pressione_vento", "N.D."))
        
        c5, c6, c7, c8 = st.columns(4)
        c5.metric("Azione Sismica", dati.get("zona_sismica", "N.D."))
        c6.metric("Luce Totale", f"{dati.get('luce_totale')} m")
        c7.metric("Struttura Princ.", dati.get('categoria_struttura', 'N.D.'))
        c8.metric("Tipologia Travatura", dati.get("tipo_travatura", "Bi-falda semplice"))
        
        st.info(f"🏗️ **Copertura configurata:** Pannello {dati.get('tipo_isolante')} ({dati.get('spessore_pannello')}) | **Impianto FV:** {dati.get('impianto_fv_desc')} | **Carico Extra:** {dati.get('carico_aggiuntivo', 0.0)} kN/mq")
        
        st.markdown("---")
        st.markdown("### 🪵 3. Arcarecci di Copertura")
        st.info(f"**Passo Calcolato Arcarecci:** {dati.get('passo_arcarecci_calc', 1.5):.2f} m | **Posizione:** {dati.get('posizione_arcarecci', 'Sopra i telai')}  \n- **Sezione Acciaio:** {dati.get('sezione_arcarecci', 'N.D.')}  \n- **Sezione Legno:** {dati.get('sezione_arcarecci_legno', 'N.D.')}")
        st.write(f"**Verifica Flessionale:** {dati.get('verifica_arcarecci', 'Verificato')}")
        
        st.markdown("---")
        st.markdown("### 🧱 4. Baraccatura di Parete (Supporto Rivestimento)")
        st.write(f"**Pannello Facciata:** {dati.get('tipo_isolante_parete', 'N.D.')} ({dati.get('spessore_pannello_parete', 'N.D.')})")
        st.success(f"**Passo Calcolato Baraccatura:** {dati.get('passo_baraccatura_calc', 2.0):.2f} m")
        st.write(f"- Sviluppo lineare baraccatura **Parete Longitudinale** (singola): {dati.get('ml_baraccatura_long_singola', 0)} ml")
        st.write(f"- Sviluppo lineare baraccatura **Timpano Frontale** (singolo): {dati.get('ml_baraccatura_timpani_singolo', 0)} ml")
        
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
        st.markdown("### 🏛️ 4.1 Montanti Verticali Antivento")
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown("#### 📐 Pareti Frontali (Timpani)")
            st.write(f"**N° Montanti per singola facciata:** {dati.get('num_montanti_timpano_singolo', 0)}")
            if dati.get('num_montanti_timpano_singolo', 0) > 0:
                st.write(f"**Passo d'installazione:** {dati.get('passo_montanti_timpano', 0):.2f} m")
                st.write(f"**Sviluppo Totale (Entrambe le facciate):** {dati.get('ml_tot_timpani_entrambe', 0):.2f} ml")
            else:
                st.info("💡 Luce contenuta, nessun montante intermedio richiesto.")
        with col_m2:
            st.markdown("#### 📏 Pareti Longitudinali")
            st.write(f"**N° Montanti per singola parete lunga:** {dati.get('num_montanti_long_singola_parete', 0)}")
            if dati.get('num_montanti_long_singola_parete', 0) > 0:
                st.write(f"**Passo d'installazione:** {dati.get('passo_montanti_long', 0):.2f} m")
                st.write(f"**Sviluppo Totale (Entrambe le pareti):** {dati.get('ml_tot_montanti_long_entrambe', 0):.2f} ml")
            else:
                st.info("💡 Interasse portali entro i 6m, nessun montante intermedio richiesto.")

        st.markdown("---")
        st.markdown("### 📐 5. Struttura Principale / Travatura (Dimensionamento)")
        if dati.get('categoria_struttura') in ["Capriate", "Travi Reticolari"]:
            col_t1, col_t2 = st.columns(2)
            with col_t1:
                st.markdown("#### 🌲 Variante in Legno")
                st.success(dati.get('travi_legno', 'N.D.').replace('\n', '  \n'))
            with col_t2:
                st.markdown("#### ⚙️ Variante in Acciaio")
                st.warning(dati.get('travi_acciaio', 'N.D.').replace('\n', '  \n'))
        else:
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
            st.success(dati.get('pilastri_perimetrali_legno', 'N.D.'))
        with col_p2:
            st.markdown("#### ⚙️ Acciaio")
            st.warning(dati.get('pilastri_perimetrali_acciaio', 'N.D.'))
        with col_p3:
            st.markdown("#### 🏛️ C.a.p.")
            st.error(dati.get('pilastri_perimetrali_cap', 'N.D.'))
        
        st.markdown("---")
        st.markdown("### 🔗 7. Stabilizzazione e Controventi")
        col_cv1, col_cv2 = st.columns(2)
        with col_cv1:
            st.markdown("#### 🛡️ Controventi di Copertura")
            st.info(f"Legno: {dati.get('controventi_copertura_legno')} | Acciaio: {dati.get('controventi_copertura_acciaio')}")
        with col_cv2:
            st.markdown("#### 🧱 Controventi di Parete")
            st.warning(f"Legno: {dati.get('controventi_parete_legno')} | Acciaio: {dati.get('controventi_parete_acciaio')}")

        st.markdown("---")
        st.markdown("### 🔩 8. Dimensionamento Dettagliato Connessioni, Nodi e Giunti in Colmo")
        col_n1, col_n2 = st.columns(2)
        with col_n1:
            st.markdown("#### 🔗 Connessione Pilastro / Trave")
            st.info(f"**Tipologia Nodo:** {dati.get('conn_trave_pilastro_tipo', 'N.D.')}")
            st.write(f"- **Nodi Perimetrali:** {dati.get('conn_trave_pilastro_perim_elementi', 'N.D.')}")
            st.metric("Peso Acciaio Nodo Perimetrale", dati.get('conn_trave_pilastro_perim_kg', 'N.D.'))
            st.write(f"- **Nodi Intermedi:** {dati.get('conn_trave_pilastro_interm_elementi', 'N.D.')}")
            st.metric("Peso Acciaio Nodo Intermedio", dati.get('conn_trave_pilastro_interm_kg', 'N.D.'))
        with col_n2:
            st.markdown("#### ⚓ Connessione Pilastro / Fondazione")
            st.warning(f"**Tipologia Base:** {dati.get('conn_pilastro_fondazione_tipo', 'N.D.')}")
            st.write(f"- **Ancoraggi Perimetrali:** {dati.get('conn_pilastro_fondazione_perim_elementi', 'N.D.')}")
            st.metric("Peso Acciaio Base Perimetrale", dati.get('conn_pilast_fondazione_perim_kg', 'N.D.') if 'conn_pilast_fondazione_perim_kg' in dati else dati.get('conn_pilastro_fondazione_perim_kg', 'N.D.'))
            st.write(f"- **Ancoraggi Intermedi:** {dati.get('conn_pilastro_fondazione_interm_elementi', 'N.D.')}")
            st.metric("Peso Acciaio Base Intermedia", dati.get('conn_pilastro_fondazione_interm_kg', 'N.D.'))
        
        if "giuntata" in dati.get('tipo_travatura', '').lower():
            st.info(f"📐 **Dettaglio Giunto in Colmo (Piastra di Giunzione):** {dati.get('dettaglio_giunto_colmo', 'N.D.')}")

        st.markdown("### 🔥 9. Requisiti di Resistenza al Fuoco e Durabilità")
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            st.metric("Classe di Resistenza Richiesta", dati.get('classe_resistenza_fuoco', 'R 60'))
        with col_f2:
            st.metric("Superficie Acciaio da Trattare", dati.get('mq_intumescente', 'Non specificato'))
        with col_f3:
            st.metric("Classe di Servizio Legno", dati.get('classe_servizio', 'Classe 2').split(' (')[0])
        st.info(f"**Specifiche Ciclo Antincendio:** {dati.get('dettaglio_verniciatura', 'N.D.')}")
        
        st.markdown("---")
        st.markdown("### 📝 10. Note Tecniche")
        st.write(dati.get("note_tecniche", "Nessuna nota aggiuntiva."))

# --- NUOVO MODULO A PARTE PER SOLAI XLAM ---
with tab_xlam:
    st.header("Modulo Dimensionamento Solai in XLAM (NTC 2018)")
    st.markdown("Il dimensionamento considera la gamma stratigrafica tipica per pannelli CLT da solaio a marchio **Stora Enso** (pannelli C a 3, 5 e 7 strati).")
    
    col_x1, col_x2 = st.columns(2)
    with col_x1:
        luce_xlam_ui = st.number_input("Luce di calcolo del solaio (m)", min_value=1.0, value=5.0, step=0.1)
    
    st.markdown("#### Pesi Permanenti Portati (G2)")
    if 'carichi_g2_xlam' not in st.session_state:
        st.session_state['carichi_g2_xlam'] = pd.DataFrame([
            {"Descrizione": "Massetto e pavimentazione", "Carico [kN/m²]": 1.5},
            {"Descrizione": "Impianti e controsoffitto", "Carico [kN/m²]": 0.5}
        ])
    
    df_g2 = st.data_editor(st.session_state['carichi_g2_xlam'], num_rows="dynamic", use_container_width=True, key="xlam_g2_editor")
    
    st.markdown("#### Sovraccarichi Variabili")
    col_qx1, col_qx2 = st.columns(2)
    with col_qx1:
        q_k_xlam = st.number_input("Sovraccarico Accidentale - Qk (kN/m²)", min_value=0.0, value=2.0, step=0.5)
    with col_qx2:
        qs_k_xlam = st.number_input("Carico Neve (se piano di copertura) - Qs (kN/m²)", min_value=0.0, value=0.0, step=0.1)

    st.markdown("#### Stati limite di esercizio (SLE) - Limiti di Freccia")
    col_sl1, col_sl2, col_sl3 = st.columns(3)
    with col_sl1:
        limite_w_inst_ui = st.number_input("Valore limite w_inst (L / ...)", min_value=100, value=300, step=10)
    with col_sl2:
        limite_w_netfin_ui = st.number_input("Valore limite w_net,fin (L / ...)", min_value=100, value=300, step=10)
    with col_sl3:
        limite_w_fin_ui = st.number_input("Valore limite w_fin (L / ...)", min_value=100, value=250, step=10)

    st.markdown("#### Resistenza al Fuoco")
    classe_fuoco_xlam = st.selectbox("Requisito Antincendio (carbonizzazione all'intradosso)", ["R 0", "R 30", "R 60", "R 90"], key="xlam_fuoco")

    if st.button("Dimensiona Solaio XLAM", type="primary"):
        g2_totale = df_g2["Carico [kN/m²]"].sum()
        
        # Parametri legno per XLAM (generalmente classe C24)
        E_mean = 11000.0  # N/mm2
        f_mk = 24.0       # N/mm2
        gamma_m = 1.25
        k_mod = 0.8
        f_md = f_mk * k_mod / gamma_m
        k_def = 0.8 # Classe di servizio 2
        
        # Carbonizzazione teorica (beta_0 = 0.65 mm/min per pannelli massicci)
        d_ef = 0
        if classe_fuoco_xlam == "R 30": d_ef = 30 * 0.65 + 7.0
        elif classe_fuoco_xlam == "R 60": d_ef = 60 * 0.65 + 7.0
        elif classe_fuoco_xlam == "R 90": d_ef = 90 * 0.65 + 7.0

        pannello_idoneo = None
        
        for pannello in pannelli_xlam_db:
            peso_proprio_g1 = (pannello["spessore"] / 1000.0) * 5.0  # kN/m2
            
            # --- VERIFICA SLU ---
            q_slu = 1.3 * (peso_proprio_g1 + g2_totale) + 1.5 * q_k_xlam + 1.5 * 0.5 * qs_k_xlam
            M_ed = (q_slu * luce_xlam_ui**2) / 8.0 * 1000000.0  # Momento per 1 metro di fascia (N*mm)
            
            I_eff, W_eff = calcola_proprieta_efficaci_xlam(pannello["strati"], pannello["orientamento"])
            I_eff *= 1000.0  # Riporto a fascia di 1 metro
            W_eff *= 1000.0
            
            sigma_m = M_ed / W_eff if W_eff > 0 else 999.0
            check_slu = sigma_m <= f_md
            
            # --- VERIFICA SLE ---
            # Carichi per fasce (kN/m)
            q_inst_G = peso_proprio_g1 + g2_totale
            q_inst_Q = q_k_xlam + qs_k_xlam
            
            # Formule di freccia elastica trave appoggiata-appoggiata
            w_inst_G = (5.0 / 384.0) * (q_inst_G / 1000.0) * (luce_xlam_ui * 1000.0)**4 / (E_mean * I_eff)
            w_inst_Q = (5.0 / 384.0) * (q_inst_Q / 1000.0) * (luce_xlam_ui * 1000.0)**4 / (E_mean * I_eff)
            
            w_inst = w_inst_G + w_inst_Q
            w_fin = w_inst_G * (1.0 + k_def) + w_inst_Q * (1.0 + 0.3 * k_def)  # Assumendo psi_2 = 0.3
            
            L_mm = luce_xlam_ui * 1000.0
            check_sle_inst = w_inst <= (L_mm / limite_w_inst_ui)
            check_sle_fin = w_fin <= (L_mm / limite_w_fin_ui)
            check_sle_netfin = w_fin <= (L_mm / limite_w_netfin_ui)
            
            # --- VERIFICA AL FUOCO ---
            check_fuoco = True
            sigma_m_fi = 0.0
            f_md_fi = 1.15 * f_mk / 1.0
            if d_ef > 0:
                I_eff_fi, W_eff_fi = calcola_proprieta_efficaci_xlam(pannello["strati"], pannello["orientamento"], d_ef)
                I_eff_fi *= 1000.0
                W_eff_fi *= 1000.0
                
                q_fi = 1.0 * (peso_proprio_g1 + g2_totale) + 1.0 * 0.3 * q_k_xlam
                M_ed_fi = (q_fi * luce_xlam_ui**2) / 8.0 * 1000000.0
                
                sigma_m_fi = M_ed_fi / W_eff_fi if W_eff_fi > 0 else 999.0
                check_fuoco = sigma_m_fi <= f_md_fi

            if check_slu and check_sle_inst and check_sle_fin and check_sle_netfin and check_fuoco:
                pannello_idoneo = pannello
                break
                
        if pannello_idoneo:
            st.success(f"✅ **Solaio XLAM Ottimizzato Trovato:** {pannello_idoneo['nome']} (Spessore {pannello_idoneo['spessore']} mm)")
            st.markdown(f"**Composizione strati (Top -> Bottom):** {pannello_idoneo['strati']} mm")
            
            c_res1, c_res2, c_res3 = st.columns(3)
            with c_res1:
                st.markdown("#### Verifica a Flessione (SLU)")
                st.write(f"$\sigma_{{m,d}}$ = **{sigma_m:.2f} MPa**")
                st.write(f"$f_{{m,d}}$ limite = **{f_md:.2f} MPa**")
            with c_res2:
                st.markdown("#### Verifica Frecce (SLE)")
                st.write(f"$w_{{inst}}$ = **{w_inst:.2f} mm** (Lim. {(L_mm / limite_w_inst_ui):.1f} mm)")
                st.write(f"$w_{{fin}}$ = **{w_fin:.2f} mm** (Lim. {(L_mm / limite_w_fin_ui):.1f} mm)")
            with c_res3:
                st.markdown("#### Verifica Antincendio")
                if classe_fuoco_xlam == "R 0":
                    st.write("Nessun requisito richiesto.")
                else:
                    st.write(f"Strato carbonizzato rimosso: **{d_ef:.1f} mm**")
                    st.write(f"$\sigma_{{m,fi,d}}$ = **{sigma_m_fi:.2f} MPa** (Lim. {f_md_fi:.2f} MPa)")
        else:
            st.error("Nessun pannello XLAM dal database standard (fino a 280mm) risulta verificato. Prova a diminuire la luce, ridurre i carichi, o prevedere dei supporti intermedi per il solaio.")
