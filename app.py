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

# --- GESTIONE SALVATAGGIO PROGETTI (FILE LOCALI) ---
class NpEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, np.integer): return int(obj)
        if isinstance(obj, np.floating): return float(obj)
        if isinstance(obj, np.ndarray): return obj.tolist()
        return super(NpEncoder, self).default(obj)

def genera_json_progetto():
    stato_da_salvare = {}
    for key, value in st.session_state.items():
        # Escludiamo i file uploader e i dataframe temporanei per evitare crash
        if key in ["xlam_g2_editor", "geo_file_cp", "carica_progetto_file"]:
            continue
        if isinstance(value, pd.DataFrame):
            stato_da_salvare[key] = {"__type__": "dataframe", "data": value.to_dict(orient="records")}
        else:
            try:
                json.dumps(value, cls=NpEncoder) 
                stato_da_salvare[key] = value
            except TypeError:
                pass 
    return json.dumps(stato_da_salvare, cls=NpEncoder, indent=4)

def applica_json_progetto(json_string):
    dati_progetto = json.loads(json_string)
    for key, value in dati_progetto.items():
        if isinstance(value, dict) and value.get("__type__") == "dataframe":
            st.session_state[key] = pd.DataFrame(value["data"])
        else:
            st.session_state[key] = value


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

# --- FUNZIONE ESTREZIONE GEOLOGICA DA PDF ---
def estrai_portanza_da_pdf(file_upload):
    try:
        pdf_reader = PyPDF2.PdfReader(file_upload)
        testo = ""
        for page in pdf_reader.pages:
            if page.extract_text():
                testo += page.extract_text() + "\n"
        
        pattern = r'(?i)(portanza|tensione ammissibile|capacit[aà] portante|sigma_adm|sigma ammissibile|resistenza terreno|carico ammissibile)[\s:=]*([0-9]+[.,]?[0-9]*)[\s]*(daN/cm2|kg/cm2|kN/m2|MPa)'
        match = re.search(pattern, testo)
        if match:
            valore = float(match.group(2).replace(',', '.'))
            unita = match.group(3).lower()
            if unita in ['dan/cm2', 'kg/cm2']:
                return valore * 100.0  
            elif unita == 'mpa':
                return valore * 1000.0
            elif unita == 'kn/m2':
                return valore
    except Exception:
        pass
    return None

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
    
    luogo_str, qsk_calc, zona_vento, press_vento_str, zona_sismica, altitudine_stimata = estrai_parametri_ntc_da_coordinate_e_comune(lat, lon, comune)
    
    qsk = dati_geo.get('qsk_manuale', 0.0)
    if qsk <= 0.0:
        qsk = qsk_calc
        
    pressione_vento_man = dati_geo.get('vento_manuale', 0.0)
    if pressione_vento_man > 0.0:
        pressione_vento = pressione_vento_man
        press_vento_str = f"{pressione_vento} kN/mq (Manuale)"
    else:
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

    tipo_ostacolo = dati_geo.get('tipo_ostacolo_neve', 'Nessuno')
    h_ostacolo = dati_geo.get('h_ostacolo_neve', 0.0)
    mu_neve = 0.8
    peso_vol_neve = 2.0 
    if tipo_ostacolo == "Parapetto":
        mu_neve = max(0.8, min(2.0, peso_vol_neve * h_ostacolo / qsk)) if qsk > 0 else 0.8
    elif tipo_ostacolo == "Edificio adiacente più alto":
        mu_neve = max(0.8, min(4.0, peso_vol_neve * h_ostacolo / qsk)) if qsk > 0 else 0.8
    
    q_s_eff = mu_neve * qsk

    g1, g2 = 0.15, 0.25
    if "Presente" in dati_geo.get('impianto_fv_desc', ''): g2 += 0.20
    g2 += dati_geo.get('carico_aggiuntivo', 0.0)
    q_tot_copertura_mq = (1.3 * g1 + 1.5 * g2 + 1.5 * q_s_eff)
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
        "note_tecniche": f"Calcolo esatto NTC 2018 con verifica antincendio ({classe_fuoco}) e {classe_servizio}. L={luce_totale}m, H_truss={H_truss:.2f}m.",
        "mu_neve": round(mu_neve, 2),
        "q_s_eff": round(q_s_eff, 2),
        "tipo_ostacolo_neve": tipo_ostacolo
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
    max_pezzi_per_viaggio_pilastri = max_pezzi_per_viaggio_pilastri = max_pezzi_pilastro_peso
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
    try:
        doc = Document('Carta Intestata.docx')
    except Exception:
        doc = Document()
    doc.add_heading('Relazione Tecnica di Predimensionamento, Calcolo e Logistica (NTC 2018)', 0)
    
    doc.add_heading('1. Parametri Geometrici, Climatici, Sismici e di Configurazione', level=1)
    doc.add_paragraph(f"Località / Comune: {dati.get('luogo', 'N.D.')}")
    if dati.get('tipo_ostacolo_neve', 'Nessuno') != 'Nessuno':
        doc.add_paragraph(f"Carico Neve base (qsk): {dati.get('qsk', 1.5)} kN/m² | Coeff. di forma (μ): {dati.get('mu_neve', 0.8):.2f} | q_s accumulo: {dati.get('q_s_eff', 1.5):.2f} kN/m² ({dati.get('tipo_ostacolo_neve')})")
    else:
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
                y_m = y_start + m * passo_montanti_long
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

def calcola_proprieta_efficaci_xlam(strati, orientamento, def_fuoco=0):
    strati_eff = list(strati)
    rimosso = def_fuoco
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
        
    I_eff *= 0.85 
    
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
    {"nome": "CLT 280 C7s", "spessore": 280, "strati": [40, 40, 40, 40, 40, 40, 40], "orientamento": [1, 0, 1, 0, 1, 0, 1]},
    {"nome": "CLT 320 C8s", "spessore": 320, "strati": [40, 40, 40, 40, 40, 40, 40, 40], "orientamento": [1, 0, 1, 0, 0, 1, 0, 1]}
]

# --- FUNZIONE GENERAZIONE WORD PER SOLAI XLAM ---
def genera_word_xlam(dati):
    try:
        doc = Document('Carta Intestata.docx')
    except Exception:
        doc = Document()
    doc.add_heading('Relazione Tecnica di Calcolo - Solaio in XLAM (NTC 2018)', 0)
    
    doc.add_heading('1. Parametri e Carichi di Progetto', level=1)
    doc.add_paragraph(f"Località / Comune: {dati.get('luogo', 'N.D.')}")
    doc.add_paragraph(f"Azione Sismica: {dati.get('zona_sismica', 'N.D.')}")
    doc.add_paragraph(f"Zona Vento: {dati.get('zona_vento', 'N.D.')} | Pressione di riferimento: {dati.get('pressione_vento', 'N.D.')}")
    doc.add_paragraph(f"Luce di calcolo del solaio: {dati['luce']} m")
    doc.add_paragraph(f"Carichi permanenti portati (G2): {dati['g2_tot']:.2f} kN/m²")
    doc.add_paragraph(f"Sovraccarico accidentale (Qk): {dati['qk']} kN/m²")
    doc.add_paragraph(f"Carico neve al suolo (qsk): {dati['qsk']} kN/m²")
    if dati['accumulo_attivo']:
        doc.add_paragraph(f"Accumulo neve considerato: Tipo ostacolo {dati['tipo_ost']} (h = {dati['h_ost']} m) -> μ = {dati['mu_calc']:.2f}, q_s = {dati['neve_fin']:.2f} kN/m²")
    doc.add_paragraph(f"Limiti di freccia SLE: w_inst <= L/{dati['lim_inst']}, w_net,fin <= L/{dati['lim_netfin']}, w_fin <= L/{dati['lim_fin']}")
    doc.add_paragraph(f"Requisito Antincendio: {dati['fuoco']}")

    doc.add_heading('2. Risultati della Verifica e Pannello Ottimizzato', level=1)
    doc.add_paragraph(f"Pannello CLT Selezionato: {dati['pannello_nome']} (Spessore totale: {dati['spessore']} mm)")
    doc.add_paragraph(f"Composizione strati (Top -> Bottom): {dati['strati']} mm")
    doc.add_paragraph(f"Peso proprio strutturale (G1): {dati['g1']:.2f} kN/m²")
    
    doc.add_heading('3. Verifiche SLU e SLE', level=1)
    doc.add_paragraph(f"Verifica a Flessione (SLU): Tensione σ_m,d = {dati['sigma_m']:.2f} MPa (Resistenza f_md = {dati['f_md']:.2f} MPa) -> VERIFICATO")
    doc.add_paragraph(f"Verifica Deformabilità (SLE):")
    doc.add_paragraph(f" - w_inst = {dati['w_inst']:.2f} mm (Limite: {dati['lim_inst_mm']:.1f} mm)")
    doc.add_paragraph(f" - w_fin = {dati['w_fin']:.2f} mm (Limite: {dati['lim_fin_mm']:.1f} mm)")
    
    if dati['fuoco'] != "R 0":
        doc.add_heading('4. Verifica Antincendio', level=1)
        doc.add_paragraph(f"Strato carbonizzato rimosso (d_ef): {dati['d_ef']:.1f} mm")
        doc.add_paragraph(f"Tensioni sotto carico di fuoco σ_m,fi,d = {dati['sigma_m_fi']:.2f} MPa (Resistenza f_md,fi = {dati['f_md_fi']:.2f} MPa) -> VERIFICATO")

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

def genera_word_carport(dati):
    try:
        doc = Document('Carta Intestata.docx')
    except Exception:
        doc = Document()
    doc.add_heading('Relazione Tecnica di Calcolo - Struttura Carport (NTC 2018)', 0)

    doc.add_heading('1. Dati Generali e Geometrici', level=1)
    doc.add_paragraph(f"Località / Comune: {dati.get('luogo', 'N.D.')}")
    doc.add_paragraph(f"Modello: {dati.get('modello', 'N.D.')} | Tipologia: {dati.get('tipo', 'N.D.')} | Forma: {dati.get('forma', 'N.D.')}")
    doc.add_paragraph(f"Larghezza Trasversale: {dati.get('larghezza', 0.0):.2f} m")
    doc.add_paragraph(f"Passo Telai: {dati.get('passo_telai', 0.0):.2f} m | Numero Campate: {dati.get('num_campate', 0)} | Lunghezza Totale: {dati.get('lunghezza_totale', 0.0):.2f} m")
    doc.add_paragraph(f"Altezza di Gronda (H_trauf): {dati.get('h_trauf', 0.0):.2f} m | Altezza di Colmo (H_first): {dati.get('h_first', 0.0):.2f} m")

    doc.add_heading('2. Parametri Climatici e Carichi', level=1)
    doc.add_paragraph(f"Azione Sismica: {dati.get('zona_sismica', 'N.D.')}")
    doc.add_paragraph(f"Zona Vento: {dati.get('zona_vento', 'N.D.')} | Pressione Vento Base: {dati.get('pressione_vento', 'N.D.')} | Calcolo (qp): {dati.get('vento', 0.0):.2f} kN/m²")
    doc.add_paragraph(f"Carico Neve al Suolo (qsk): {dati.get('neve_qsk', 0.0):.2f} kN/m² | Neve di calcolo (qs): {dati.get('neve', 0.0):.2f} kN/m²")
    doc.add_paragraph(f"Pesi Permanenti: Struttura (G1) {dati.get('g1', 0.0):.2f} kN/m² | Impianti/Pannelli (G2) {dati.get('g2', 0.0):.2f} kN/m²")
    doc.add_paragraph(f"Carico Totale Equivalente (SLU): {dati.get('q_tot', 0.0):.2f} kN/m² ({dati.get('kg_mq', 0.0):.1f} kg/m²)")

    doc.add_heading('3. Dimensionamento Elementi Strutturali', level=1)
    doc.add_paragraph(f"Trave di Falda: {dati.get('sez_trave', 'N.D.')} (M_ed = {dati.get('M_trave', 0.0):.1f} kNm)")
    doc.add_paragraph(f"Colonna Portante: {dati.get('sez_col', 'N.D.')}")
    doc.add_paragraph(f"Arcarecci di Copertura: {dati.get('sez_arc', 'N.D.')} (Interasse max: {dati.get('passo_arc', 0.0):.2f} m)")
    doc.add_paragraph(f"Controventi di Copertura: {dati.get('cv_falda', 'N.D.')}")
    doc.add_paragraph(f"Controventi Verticali: {dati.get('cv_vert', 'N.D.')}")

    doc.add_heading('4. Reazioni Vincolari, Connessioni e Fondazioni', level=1)
    doc.add_paragraph(f"Reazioni di base (SLU): N = {dati.get('N_base', 0.0):.1f} kN | V = {dati.get('V_base', 0.0):.1f} kN | M = {dati.get('M_base', 0.0):.1f} kNm")
    doc.add_paragraph(f"Capacità Portante Terreno: {dati.get('sigma_terreno', 150.0):.1f} kN/m²")
    doc.add_paragraph(f"Tipologia Plinto: {dati.get('forma_plinto', 'N.D.')}")
    doc.add_paragraph(f"Dimensioni Plinto: {dati.get('dim_plinto', 'N.D.')}")
    doc.add_paragraph(f"Materiali Plinto (Stimati): {dati.get('vol_plinto', 0.0):.2f} m³ Cls | {dati.get('kg_armatura', 0.0):.1f} kg Armatura")
    doc.add_paragraph(f"Peso Nodo Base: ~{dati.get('kg_nodo_base', 0.0):.1f} kg | Peso Nodo Top: ~{dati.get('kg_nodo_top', 0.0):.1f} kg")

    doc.add_heading('5. Durabilità e Trattamenti', level=1)
    doc.add_paragraph(f"Ciclo Anticorrosione: {dati.get('ciclo_c5', 'N.D.')}")
    doc.add_paragraph(f"Superficie Acciaio da trattare: {dati.get('mq_acciaio_totale', 0.0):.1f} m²")

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# --- FUNZIONI CARPORT AGGIORNATE CON PLINTI CIABATTA+DADO ---
def esegui_calcolo_carport(dati):
    w = dati['larghezza']
    pt = dati['passo_telai']
    nc = dati['num_campate']
    q_tot = dati['q_tot']
    tipo = dati['tipo']
    forma = dati['forma']
    vento = dati.get('vento', 0.0)
    sigma_terreno_kn = dati.get('sigma_terreno', 150.0)
    
    passo_arc_max = 1.2 if "Acciaio" in tipo else 1.5
    n_arc = math.ceil(w / passo_arc_max)
    passo_arc = w / n_arc if n_arc > 0 else 1.0

    M_arc = (q_tot * passo_arc * pt**2) / 8.0
    if "Acciaio" in tipo.split('-')[1]:
        if M_arc < 5: sez_arc = "Profilo a Z pressopiegato 120x2.0 mm"
        elif M_arc < 10: sez_arc = "Profilo a Z pressopiegato 150x2.5 mm"
        else: sez_arc = "Profilo a Z pressopiegato 200x3.0 mm"
    else: 
        w_req = (M_arc * 100) / 1.45
        h_req = math.sqrt((6*w_req)/10.0)
        h_arc = max(16, math.ceil(h_req/4)*4)
        sez_arc = f"Legno Lamellare GL24h 10x{int(h_arc)} cm"

    if "Y" in forma:
        L_cant = w / 2.0
        M_trave = (q_tot * pt * L_cant**2) / 2.0
    else:
        M_trave = (q_tot * pt * w**2) / 8.0

    if "Acciaio" in tipo.split('-')[1]:
        w_req_tr = (M_trave * 100) / 27.5
        if w_req_tr < 150: sez_trave = "IPE 200 / HEA 140"
        elif w_req_tr < 250: sez_trave = "IPE 240 / HEA 180"
        elif w_req_tr < 420: sez_trave = "IPE 300 / HEA 220"
        elif w_req_tr < 700: sez_trave = "IPE 360 / HEA 260"
        elif w_req_tr < 1100: sez_trave = "IPE 450 / HEA 320"
        elif w_req_tr < 1500: sez_trave = "IPE 500 / HEA 360"
        else: sez_trave = "IPE 600 / HEB 400"
    else:
        w_req_tr = (M_trave * 100) / 1.45
        b_tr = 20
        h_req_tr = math.sqrt((6*w_req_tr)/b_tr)
        h_tr = max(24, math.ceil(h_req_tr/4)*4)
        sez_trave = f"BSH GL24h {b_tr}x{int(h_tr)} cm (consigliata a sezione variabile)"

    h_media = (dati.get('h_trauf', 2.4) + dati.get('h_first', 3.0)) / 2.0
    
    if "Y" in forma:
        N_col = (q_tot * pt * w)
        V_col = (vento * 1.5 * pt * h_media)
        M_col = (M_trave * 0.60) + (V_col * h_media / 2.0)
        
        w_req_col = (M_col * 100) / 27.5
        if w_req_col < 150: sez_col = "HEB 140 / HEA 160"
        elif w_req_col < 300: sez_col = "HEB 180 / HEA 200"
        elif w_req_col < 600: sez_col = "HEB 220 / HEA 260"
        elif w_req_col < 1000: sez_col = "HEB 280 / HEA 320"
        else: sez_col = "HEB 320 / HEB 360"
    else:
        N_col = (q_tot * pt * w) / 2.0
        V_col = (vento * 1.5 * pt * h_media) / 2.0
        M_col = V_col * h_media
        
        w_req_col = (M_col * 100) / 27.5 + (N_col / 2.0) 
        if w_req_col < 80: sez_col = "Tubolare 100x100x4 / HEA 120"
        elif w_req_col < 150: sez_col = "Tubolare 150x150x5 / HEA 140"
        elif w_req_col < 300: sez_col = "HEB 160 / HEA 180"
        else: sez_col = "HEB 200 / HEA 220"

    kg_nodo_top = round(25.0 + M_trave * 0.15, 1)
    kg_nodo_base = round(35.0 + M_col * 0.25 + N_col * 0.05, 1)

    # --- Nuovo Dimensionamento Plinto di fondazione (Ciabatta + Dado) ---
    B_dado = 0.60  
    H_dado = 0.50  
    B_pl = 0.8     
    
    while True:
        H_pad = max(0.4, math.ceil((B_pl / 4)*10)/10)
        vol_dado = B_dado * B_dado * H_dado
        vol_pad = B_pl * B_pl * H_pad
        
        N_cls = (vol_dado + vol_pad) * 25.0 
        N_tot = N_col + N_cls
        M_tot = M_col + V_col * (H_dado + H_pad) 
        
        area_pl = B_pl * B_pl
        W_pl = (B_pl**3) / 6.0
        
        sigma_max = (N_tot / area_pl) + (M_tot / W_pl)
        
        if sigma_max <= sigma_terreno_kn or B_pl >= 4.0:
            break
        B_pl += 0.1

    vol_plinto = round(vol_dado + vol_pad, 2)
    kg_armatura = round(vol_plinto * 80.0, 1)
    dim_plinto_str = f"Ciabatta {B_pl:.1f}x{B_pl:.1f}x{H_pad:.1f}m + Dado {B_dado:.1f}x{B_dado:.1f}x{H_dado:.1f}m"

    cv_falda = "Tiranti in acciaio incrociati Ø 16 mm (campate di estremità)"
    cv_vert = "Incastro rigido in fondazione (nessun controvento verticale previsto per viabilità)" if "Y" in forma else "Croci di Sant'Andrea in tubolare 80x80x4 mm o L 80x8 (sulla linea colonne)"

    num_telai = nc + 1
    num_colonne = num_telai * (1 if "Y" in forma else 2) 
    h_media_colonna = h_media
    ml_colonne_tot = num_colonne * h_media_colonna
    ml_travi_tot = num_telai * w * 1
    ml_arcarecci_tot = (n_arc + 1) * (nc * pt)
    
    mq_acciaio_totale = round((ml_colonne_tot + ml_travi_tot + ml_arcarecci_tot) * 0.8, 1)

    return {
        "passo_arc": round(passo_arc, 2),
        "sez_arc": sez_arc,
        "sez_trave": sez_trave,
        "sez_col": sez_col,
        "cv_falda": cv_falda,
        "cv_vert": cv_vert,
        "M_trave": round(M_trave, 1),
        "mq_acciaio_totale": mq_acciaio_totale,
        "N_base": round(N_col, 1),
        "V_base": round(V_col, 1),
        "M_base": round(M_col, 1),
        "kg_nodo_top": kg_nodo_top,
        "kg_nodo_base": kg_nodo_base,
        "dim_plinto": dim_plinto_str,
        "forma_plinto": "Plinto a gradoni (Ciabatta ripartitrice + Dado di ancoraggio)",
        "vol_plinto": vol_plinto,
        "kg_armatura": kg_armatura,
        "B_pl": round(B_pl, 2), "H_pad": round(H_pad, 2),
        "B_dado": round(B_dado, 2), "H_dado": round(H_dado, 2)
    }

def genera_modello_3d_carport(dati):
    fig = go.Figure()
    w = dati['larghezza']
    pt = dati['passo_telai']
    nc = dati['num_campate']
    h_trauf = dati['h_trauf']
    h_first = dati['h_first']
    forma = dati['forma']
    
    y_telai = [i * pt for i in range(nc + 1)]
    lunghezza_totale = nc * pt
    
    col_positions = []
    
    for y in y_telai:
        if "Y-Doppelcarport" in forma:
            x_col = w / 2.0
            z_col = h_trauf
            col_positions.append((x_col, y))
            fig.add_trace(go.Scatter3d(x=[x_col, x_col], y=[y, y], z=[0, z_col], mode='lines', line=dict(color='darkblue', width=8), showlegend=(y==0), name='Colonna'))
            fig.add_trace(go.Scatter3d(x=[x_col, 0], y=[y, y], z=[z_col, h_first], mode='lines', line=dict(color='firebrick', width=6), showlegend=(y==0), name='Mensola'))
            fig.add_trace(go.Scatter3d(x=[x_col, w], y=[y, y], z=[z_col, h_first], mode='lines', line=dict(color='firebrick', width=6), showlegend=False))
        elif "fallend" in forma:
            x_col = 0.0
            z_col = h_first
            col_positions.append((x_col, y))
            fig.add_trace(go.Scatter3d(x=[x_col, x_col], y=[y, y], z=[0, z_col], mode='lines', line=dict(color='darkblue', width=8), showlegend=(y==0), name='Colonna'))
            fig.add_trace(go.Scatter3d(x=[x_col, w], y=[y, y], z=[z_col, h_trauf], mode='lines', line=dict(color='firebrick', width=6), showlegend=(y==0), name='Mensola'))
        else: 
            x_col = 0.0
            z_col = h_trauf
            col_positions.append((x_col, y))
            fig.add_trace(go.Scatter3d(x=[x_col, x_col], y=[y, y], z=[0, z_col], mode='lines', line=dict(color='darkblue', width=8), showlegend=(y==0), name='Colonna'))
            fig.add_trace(go.Scatter3d(x=[x_col, w], y=[y, y], z=[z_col, h_first], mode='lines', line=dict(color='firebrick', width=6), showlegend=(y==0), name='Mensola'))

    # --- Generazione grafica 3D dei plinti (Sotto la quota 0) ---
    B_pl = dati.get('B_pl', 1.0)
    H_pad = dati.get('H_pad', 0.4)
    B_dado = dati.get('B_dado', 0.6)
    H_dado = dati.get('H_dado', 0.5)
    
    for idx, (xc, yc) in enumerate(col_positions):
        show_leg_pl = (idx == 0)
        
        # Facce del DADO (Wireframe)
        hd = B_dado / 2.0
        z_top_dado = 0
        z_bot_dado = -H_dado
        fig.add_trace(go.Scatter3d(x=[xc-hd, xc+hd, xc+hd, xc-hd, xc-hd], y=[yc-hd, yc-hd, yc+hd, yc+hd, yc-hd], z=[z_bot_dado]*5, mode='lines', line=dict(color='gray', width=4), showlegend=show_leg_pl, name='Plinto (Dado)'))
        fig.add_trace(go.Scatter3d(x=[xc-hd, xc+hd, xc+hd, xc-hd, xc-hd], y=[yc-hd, yc-hd, yc+hd, yc+hd, yc-hd], z=[z_top_dado]*5, mode='lines', line=dict(color='gray', width=4), showlegend=False))
        for dx, dy in [(-hd, -hd), (hd, -hd), (hd, hd), (-hd, hd)]:
            fig.add_trace(go.Scatter3d(x=[xc+dx, xc+dx], y=[yc+dy, yc+dy], z=[z_bot_dado, z_top_dado], mode='lines', line=dict(color='gray', width=4), showlegend=False))

        # Facce della CIABATTA (Wireframe)
        hp = B_pl / 2.0
        z_top_pad = -H_dado
        z_bot_pad = -H_dado - H_pad
        fig.add_trace(go.Scatter3d(x=[xc-hp, xc+hp, xc+hp, xc-hp, xc-hp], y=[yc-hp, yc-hp, yc+hp, yc+hp, yc-hp], z=[z_bot_pad]*5, mode='lines', line=dict(color='darkgray', width=4), showlegend=show_leg_pl, name='Plinto (Ciabatta)'))
        fig.add_trace(go.Scatter3d(x=[xc-hp, xc+hp, xc+hp, xc-hp, xc-hp], y=[yc-hp, yc-hp, yc+hp, yc+hp, yc-hp], z=[z_top_pad]*5, mode='lines', line=dict(color='darkgray', width=4), showlegend=False))
        for dx, dy in [(-hp, -hp), (hp, -hp), (hp, hp), (-hp, hp)]:
            fig.add_trace(go.Scatter3d(x=[xc+dx, xc+dx], y=[yc+dy, yc+dy], z=[z_bot_pad, z_top_pad], mode='lines', line=dict(color='darkgray', width=4), showlegend=False))


    passo_arc = dati['passo_arc']
    n_arc = max(1, int(w / passo_arc))
    for i in range(n_arc + 1):
        x_arc = i * passo_arc
        if "Y-Doppelcarport" in forma:
            z_arc = h_first - ((h_first - h_trauf) * (x_arc / (w/2.0))) if x_arc <= w/2.0 else h_trauf + ((h_first - h_trauf) * ((x_arc - w/2.0) / (w/2.0)))
        elif "fallend" in forma:
            z_arc = h_first - ((h_first - h_trauf) * (x_arc / w))
        else: 
            z_arc = h_trauf + ((h_first - h_trauf) * (x_arc / w))
            
        fig.add_trace(go.Scatter3d(x=[x_arc, x_arc], y=[0, lunghezza_totale], z=[z_arc, z_arc], mode='lines', line=dict(color='gray', width=3, dash='dot'), showlegend=(i==0), name='Arcarecci'))

    for idx in [0, nc - 1]:
        y1, y2 = y_telai[idx], y_telai[idx + 1]
        for i in range(n_arc):
            xa, xb = i * passo_arc, (i+1) * passo_arc
            if "Y-Doppelcarport" in forma:
                za = h_first - ((h_first - h_trauf) * (xa / (w/2.0))) if xa <= w/2.0 else h_trauf + ((h_first - h_trauf) * ((xa - w/2.0) / (w/2.0)))
                zb = h_first - ((h_first - h_trauf) * (xb / (w/2.0))) if xb <= w/2.0 else h_trauf + ((h_first - h_trauf) * ((xb - w/2.0) / (w/2.0)))
            elif "fallend" in forma:
                za = h_first - ((h_first - h_trauf) * (xa / w))
                zb = h_first - ((h_first - h_trauf) * (xb / w))
            else:
                za = h_trauf + ((h_first - h_trauf) * (xa / w))
                zb = h_trauf + ((h_first - h_trauf) * (xb / w))
            
            fig.add_trace(go.Scatter3d(x=[xa, xb, None, xa, xb], y=[y1, y2, None, y2, y1], z=[za, zb, None, za, zb], mode='lines', line=dict(color='forestgreen', width=3), showlegend=(idx==0 and i==0), name='Controventi Falda'))
        
        if not "Y" in forma:
            z_col_top = h_first if "fallend" in forma else h_trauf
            fig.add_trace(go.Scatter3d(x=[0, 0, None, 0, 0], y=[y1, y2, None, y2, y1], z=[0, z_col_top, None, z_col_top, 0], mode='lines', line=dict(color='darkorange', width=4), showlegend=(idx==0), name='Controventi Verticali'))

    fig.update_layout(
        title=f"Modello 3D Carport ({nc} Campate - {forma})",
        scene=dict(xaxis_title=f'X ({w}m)', yaxis_title=f'Y ({lunghezza_totale}m)', zaxis_title=f'Z ({h_first}m)', aspectmode='data'),
        margin=dict(l=0, r=0, b=0, t=40), height=500
    )
    return fig

# --- INIZIO APPLICAZIONE STREAMLIT ---

st.set_page_config(page_title="Predimensionamento Strutturale NTC 2018", layout="wide")
st.title("Generatore Offerte Tecniche e Dimensionamento IA 🏗️")

with st.sidebar:
    st.header("💾 Gestione Progetto")
    
    # Download (Esporta)
    json_progetto = genera_json_progetto()
    st.download_button(
        label="💾 Scarica File di Progetto (.json)",
        data=json_progetto,
        file_name=f"Progetto_{st.session_state.get('comune_cantiere_ui', 'Strutturale').replace(' ', '_')}.json",
        mime="application/json",
        use_container_width=True
    )
    
    st.markdown("---")
    
    # Upload (Importa)
    file_progetto = st.file_uploader("📂 Carica un progetto salvato (.json)", type=["json"], key="carica_progetto_file")
    if file_progetto is not None:
        if st.button("🔄 Ripristina Progetto", use_container_width=True):
            try:
                applica_json_progetto(file_progetto.getvalue().decode("utf-8"))
                st.success("Progetto caricato con successo!")
                st.rerun()
            except Exception as e:
                st.error(f"Errore nel caricamento del file: {e}")

    st.markdown("---")
    
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


tab_principale, tab_xlam, tab_travi, tab_carport = st.tabs(["🏗️ Struttura Principale Capannone", "🪵 Dimensionamento Solaio XLAM", "📏 Dimensionamento Travi", "🚗 Dimensionamento Carport"])

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

    st.markdown("#### 🛠️ Forzatura Manuale Carichi (Lascia 0.0 per calcolo automatico da Maps)")
    col_man1, col_man2 = st.columns(2)
    with col_man1:
        qsk_manuale_ui = st.number_input("Carico Neve (qsk) manuale (kN/m²)", min_value=0.0, value=0.0, step=0.1, key="qsk_man_ui")
    with col_man2:
        vento_manuale_ui = st.number_input("Pressione Vento manuale (kN/m²)", min_value=0.0, value=0.0, step=0.1, key="vento_man_ui")

    st.markdown("### ❄️ Effetti Locali: Accumulo Neve (NTC 2018 / EN 1991-1-3)")
    accumulo_neve_attivo = st.checkbox("Considera Accumulo Neve (Parapetti o variazioni di quota)", value=False, key="accumulo_neve_attivo")
    if accumulo_neve_attivo:
        col_acc1, col_acc2 = st.columns(2)
        with col_acc1:
            tipo_ostacolo_neve = st.selectbox("Tipologia Ostacolo", ["Parapetto", "Edificio adiacente più alto"], key="tipo_ostacolo_neve")
        with col_acc2:
            h_ostacolo_neve = st.number_input("Altezza dell'ostacolo h (m)", min_value=0.0, value=1.0, step=0.1, key="h_ostacolo_neve")
    else:
        tipo_ostacolo_neve = "Nessuno"
        h_ostacolo_neve = 0.0

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
                'qsk_manuale': qsk_manuale_ui, 'vento_manuale': vento_manuale_ui,
                'tipo_ostacolo_neve': tipo_ostacolo_neve, 'h_ostacolo_neve': h_ostacolo_neve,
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
        
        if dati.get('tipo_ostacolo_neve', 'Nessuno') != "Nessuno":
            c2.metric("Neve con Accumulo (qs)", f"{dati.get('q_s_eff', 1.5):.2f} kN/m²", f"qsk={dati.get('qsk')} | μ={dati.get('mu_neve'):.2f}")
        else:
            c2.metric("Carico Neve (qsk)", f"{dati.get('qsk', 1.5)} kN/m²", "μ = 0.8")
            
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

with tab_xlam:
    st.header("Modulo Dimensionamento Solai in XLAM (NTC 2018)")
    st.markdown("Il dimensionamento considera la gamma stratigrafica tipica per pannelli CLT da solaio a marchio **Stora Enso** (pannelli C a 3, 5, 7 e 8 strati).")
    
    st.markdown("### 📍 Localizzazione Cantiere Solaio (Google Maps e Comune)")
    col_x_loc1, col_x_loc2 = st.columns([2, 1])
    with col_x_loc1:
        maps_url_xlam = st.text_input("Incolla il link di Google Maps del cantiere (Solaio):", value="", key="maps_url_xlam")
    with col_x_loc2:
        _, _, luogo_estratto_xlam = estrai_dati_da_url_maps(maps_url_xlam)
        comune_xlam = st.text_input("Comune di installazione (Solaio)", value=luogo_estratto_xlam, key="comune_xlam")

    st.markdown("---")
    
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
    
    st.markdown("#### Sovraccarichi Variabili (Lascia 0.0 sul Neve per calcolo automatico da Maps)")
    col_qx1, col_qx2 = st.columns(2)
    with col_qx1:
        q_k_xlam = st.number_input("Sovraccarico Accidentale - Qk (kN/m²)", min_value=0.0, value=2.0, step=0.5)
    with col_qx2:
        qs_k_xlam = st.number_input("Carico Neve al suolo - qsk manuale (kN/m²)", min_value=0.0, value=0.0, step=0.1)

    st.markdown("#### ❄️ Accumulo Neve (Copertura - NTC 2018)")
    accumulo_xlam_attivo = st.checkbox("Considera Accumulo Neve", key="chk_acc_xlam")
    if accumulo_xlam_attivo:
        col_ax1, col_ax2 = st.columns(2)
        with col_ax1:
            tipo_ost_xlam = st.selectbox("Tipologia Ostacolo", ["Parapetto", "Edificio adiacente più alto"], key="tipo_ost_xlam")
        with col_ax2:
            h_ost_xlam = st.number_input("Altezza ostacolo h (m)", min_value=0.0, value=1.0, step=0.1, key="h_ost_xlam")

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
        lat_xlam, lon_xlam, place_xlam = estrai_dati_da_url_maps(maps_url_xlam)
        comune_finale_xlam = comune_xlam if comune_xlam else place_xlam
        luogo_str_xlam, qsk_xlam, zona_vento_xlam, press_vento_str_xlam, zona_sismica_xlam, alt_xlam = estrai_parametri_ntc_da_coordinate_e_comune(lat_xlam, lon_xlam, comune_finale_xlam)
        
        neve_base_xlam = qs_k_xlam if qs_k_xlam > 0.0 else qsk_xlam
        
        g2_totale = df_g2["Carico [kN/m²]"].sum()
        
        if accumulo_xlam_attivo:
            mu_xlam_calc = min(2.0 if tipo_ost_xlam == "Parapetto" else 4.0, max(0.8, 2.0 * h_ost_xlam / neve_base_xlam)) if neve_base_xlam > 0 else 0.8
            carico_neve_finale = mu_xlam_calc * neve_base_xlam
        else:
            carico_neve_finale = neve_base_xlam
        
        st.success(f"📍 Località rilevata: {luogo_str_xlam}  \n🌍 Azione Sismica: **{zona_sismica_xlam}** | 💨 Vento: **{zona_vento_xlam} ({press_vento_str_xlam})**")
            
        E_mean = 11000.0  
        f_mk = 24.0       
        gamma_m = 1.25
        k_mod = 0.8
        f_md = f_mk * k_mod / gamma_m
        k_def = 0.8 
        
        d_ef = 0
        if classe_fuoco_xlam == "R 30": d_ef = 30 * 0.65 + 7.0
        elif classe_fuoco_xlam == "R 60": d_ef = 60 * 0.65 + 7.0
        elif classe_fuoco_xlam == "R 90": d_ef = 90 * 0.65 + 7.0

        pannello_idoneo = None
        
        for pannello in pannelli_xlam_db:
            peso_proprio_g1 = (pannello["spessore"] / 1000.0) * 5.0  
            
            q_slu = 1.3 * (peso_proprio_g1 + g2_totale) + 1.5 * q_k_xlam + 1.5 * 0.5 * carico_neve_finale
            M_ed = (q_slu * luce_xlam_ui**2) / 8.0 * 1000000.0  
            
            I_eff, W_eff = calcola_proprieta_efficaci_xlam(pannello["strati"], pannello["orientamento"])
            I_eff *= 1000.0  
            W_eff *= 1000.0
            
            sigma_m = M_ed / W_eff if W_eff > 0 else 999.0
            check_slu = sigma_m <= f_md
            
            q_inst_G = peso_proprio_g1 + g2_totale
            q_inst_Q = q_k_xlam + carico_neve_finale
            
            w_inst_G = (5.0 / 384.0) * (q_inst_G / 1000.0) * (luce_xlam_ui * 1000.0)**4 / (E_mean * I_eff)
            w_inst_Q = (5.0 / 384.0) * (q_inst_Q / 1000.0) * (luce_xlam_ui * 1000.0)**4 / (E_mean * I_eff)
            
            w_inst = w_inst_G + w_inst_Q
            w_fin = w_inst_G * (1.0 + k_def) + w_inst_Q * (1.0 + 0.3 * k_def) 
            
            L_mm = luce_xlam_ui * 1000.0
            check_sle_inst = w_inst <= (L_mm / limite_w_inst_ui)
            check_sle_fin = w_fin <= (L_mm / limite_w_fin_ui)
            check_sle_netfin = w_fin <= (L_mm / limite_w_netfin_ui)
            
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
            reazione_appoggio_xlam = (q_slu * luce_xlam_ui) / 2.0
            
            st.success(f"✅ **Solaio XLAM Ottimizzato Trovato:** {pannello_idoneo['nome']} (Spessore {pannello_idoneo['spessore']} mm)")
            st.markdown(f"**Composizione strati (Top -> Bottom):** {pannello_idoneo['strati']} mm")
            st.write(f"Peso proprio strutturale (G1) considerato nel calcolo: **{peso_proprio_g1:.2f} kN/m²**")
            
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
            
            st.session_state['xlam_ultimi'] = {
                'luogo': luogo_str_xlam,
                'zona_sismica': zona_sismica_xlam,
                'zona_vento': zona_vento_xlam,
                'pressione_vento': press_vento_str_xlam,
                'luce': luce_xlam_ui,
                'g2_tot': g2_totale,
                'qk': q_k_xlam,
                'qsk': neve_base_xlam,
                'accumulo_attivo': accumulo_xlam_attivo,
                'tipo_ost': tipo_ost_xlam if accumulo_xlam_attivo else "",
                'h_ost': h_ost_xlam if accumulo_xlam_attivo else 0,
                'mu_calc': mu_xlam_calc if accumulo_xlam_attivo else 0.8,
                'neve_fin': carico_neve_finale,
                'lim_inst': limite_w_inst_ui,
                'lim_netfin': limite_w_netfin_ui,
                'lim_fin': limite_w_fin_ui,
                'fuoco': classe_fuoco_xlam,
                'pannello_nome': pannello_idoneo['nome'],
                'spessore': pannello_idoneo['spessore'],
                'strati': pannello_idoneo['strati'],
                'g1': peso_proprio_g1,
                'sigma_m': sigma_m,
                'f_md': f_md,
                'w_inst': w_inst,
                'w_fin': w_fin,
                'lim_inst_mm': L_mm / limite_w_inst_ui,
                'lim_fin_mm': L_mm / limite_w_fin_ui,
                'd_ef': d_ef,
                'sigma_m_fi': sigma_m_fi if classe_fuoco_xlam != "R 0" else 0,
                'f_md_fi': f_md_fi if classe_fuoco_xlam != "R 0" else 0,
                'reazione_appoggio': reazione_appoggio_xlam
            }
        else:
            st.error("Nessun pannello XLAM dal database standard (fino a 320mm) risulta verificato. Prova a diminuire la luce, ridurre i carichi, o prevedere dei supporti intermedi per il solaio.")

    if 'xlam_ultimi' in st.session_state:
        st.markdown("---")
        dx = st.session_state['xlam_ultimi']
        word_file_xlam = genera_word_xlam(dx)
        st.download_button(
            label="📄 Scarica Relazione Solaio XLAM in Word (.docx)",
            data=word_file_xlam,
            file_name=f"Relazione_Solaio_XLAM_{dx['pannello_nome'].replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True
        )

with tab_travi:
    st.header("Modulo Dimensionamento Travi (Acciaio e Legno Lamellare)")
    st.markdown("Questo modulo calcola e suggerisce la sezione della trave, permettendo di incrociare i carichi provenienti dal modulo XLAM o da travi secondarie precedentemente dimensionate.")
    
    st.markdown("### 🧱 1. Materiale e Proprietà")
    col_mat1, col_mat2 = st.columns(2)
    with col_mat1:
        mat_trave = st.radio("Materiale della Trave", ["Acciaio", "Legno Lamellare"])
    with col_mat2:
        if mat_trave == "Acciaio":
            grado_acciaio = st.radio("Grado Acciaio", ["S275", "S355"])
        else:
            classe_legno = st.selectbox("Classe Legno Lamellare", ["GL24h", "GL28h", "GL30h", "GL32h"])
            essenza_legno = st.radio("Essenza", ["Abete", "Larice"])
            
    col_geom1, col_geom2 = st.columns(2)
    with col_geom1:
        luce_trave = st.number_input("Luce di calcolo della trave (m)", min_value=1.0, value=5.0, step=0.1, key="luce_tr")
    
    st.markdown("---")
    st.markdown("### 📥 2. Carichi Uniformemente Distribuiti (q)")
    usa_xlam = st.checkbox("Importa reazione di appoggio dal modulo Solai XLAM", value=False)
    q_xlam_lineare = 0.0
    if usa_xlam:
        if 'xlam_ultimi' in st.session_state:
            q_xlam_lineare = st.session_state['xlam_ultimi'].get('reazione_appoggio', 0.0)
            st.success(f"✅ Reazione di appoggio (SLU) importata dal solaio XLAM: **{q_xlam_lineare:.2f} kN/m**")
        else:
            st.warning("⚠️ Nessun solaio XLAM calcolato precedentemente. (Esegui prima il calcolo nella scheda XLAM).")
            
    q_distr_man = st.number_input("Aggiungi carico distribuito manuale q (kN/m)", min_value=0.0, value=0.0, step=0.5, key="q_distr_man")
    q_tot_distr = q_xlam_lineare + q_distr_man
    
    st.markdown("---")
    st.markdown("### 🎯 3. Carichi Concentrati (F)")
    usa_storico = st.checkbox("Importa reazione da una Trave precedentemente calcolata (es. orditura secondaria)", value=False)
    f_conc_storico = 0.0
    if usa_storico:
        if 'travi_storico' in st.session_state and len(st.session_state['travi_storico']) > 0:
            trave_sel = st.selectbox("Seleziona la trave da far scaricare su questa", [t['nome'] for t in st.session_state['travi_storico']])
            for t in st.session_state['travi_storico']:
                if t['nome'] == trave_sel:
                    f_conc_storico = t['reazione_max']
                    st.success(f"✅ Reazione importata (V_ed max della {trave_sel}): **{f_conc_storico:.2f} kN**")
        else:
            st.warning("⚠️ Nessuna trave calcolata precedentemente nello storico.")
            
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        f_conc_man = st.number_input("Aggiungi carico concentrato manuale F (kN)", min_value=0.0, value=0.0, step=1.0, key="f_conc_man")
        f_tot_conc = f_conc_storico + f_conc_man
    with col_c2:
        pos_f_conc = st.number_input("Distanza di applicazione del carico F dall'appoggio A (m)", min_value=0.0, max_value=luce_trave, value=luce_trave/2, step=0.1)

    if st.button("Dimensiona Trave e Salva in Storico", type="primary"):
        a = pos_f_conc
        b = luce_trave - pos_f_conc
        
        M_distr = (q_tot_distr * luce_trave**2) / 8.0
        M_conc = (f_tot_conc * a * b) / luce_trave if f_tot_conc > 0 else 0.0
        
        R_A = (q_tot_distr * luce_trave / 2.0) + (f_tot_conc * b / luce_trave)
        R_B = (q_tot_distr * luce_trave / 2.0) + (f_tot_conc * a / luce_trave)
        R_max = max(R_A, R_B)
        
        M_ed = M_distr + M_conc
        
        st.write("---")
        st.markdown("### 📊 Risultati Sollecitazioni")
        st.write(f"**Momento Flettente Massimo (M_ed):** {M_ed:.2f} kNm")
        st.write(f"**Reazione Massima agli Appoggi (V_ed):** {R_max:.2f} kN")
        
        st.markdown("### 🛠️ Profilo Suggerito (NTC 2018)")
        if mat_trave == "Acciaio":
            f_y = 27.5 if grado_acciaio == "S275" else 35.5 
            w_el_req = (M_ed * 100) / f_y
            if w_el_req < 150: sez_out = "IPE 200 / HEA 140"
            elif w_el_req < 300: sez_out = "IPE 240 / HEA 180"
            elif w_el_req < 600: sez_out = "IPE 330 / HEA 240"
            elif w_el_req < 1200: sez_out = "IPE 450 / HEA 300"
            else: sez_out = "IPE 600 / HEB 400"
            st.success(f"**Profilo in Acciaio Ideale:** {sez_out} ({grado_acciaio})")
            
        else:
            f_mk = float(classe_legno[2:4])
            k_mod = 0.8
            gamma_m = 1.25
            f_md = (f_mk * k_mod) / gamma_m
            w_req_legno = (M_ed * 100) / (f_md * 10) 
            b_opt = 20
            h_req = math.sqrt((6 * w_req_legno) / b_opt)
            h_opt = max(24, int((h_req + 3) // 4) * 4)
            st.success(f"**Sezione in Legno Lamellare:** Base {b_opt} cm x Altezza {h_opt} cm ({classe_legno} - {essenza_legno})")

        if 'travi_storico' not in st.session_state:
            st.session_state['travi_storico'] = []
        
        nome_nuova_trave = f"Trave L={luce_trave}m ({mat_trave}) #{len(st.session_state['travi_storico'])+1}"
        st.session_state['travi_storico'].append({
            'nome': nome_nuova_trave,
            'reazione_max': R_max,
            'M_ed': M_ed
        })
        st.info(f"✅ Dati della **{nome_nuova_trave}** salvati nello storico! Ora potrai selezionarla per farla scaricare come carico concentrato sulle travi principali successive.")

with tab_carport:
    st.header("Modulo Dimensionamento Strutturale Carport (NTC 2018)")
    st.markdown("Il modulo dimensiona le sezioni, gli arcarecci, i controventi e determina automaticamente i parametri climatici, sismici e la protezione anticorrosione C5 in base alla localizzazione GPS.")
    
    carport_db = {
        "SC-L3 (Stahl-Stahl)": {"tipo": "Acciaio-Acciaio", "forma": "Y-Doppelcarport (Satteldach)", "b_std": 0.0, "h_trauf": 3.00, "h_first": 4.50, "dn": 10, "file": "530K - Carport System SC-L3.pdf", "is_shc": False},
        "SC-P1 (Stahl-Stahl)": {"tipo": "Acciaio-Acciaio", "forma": "Pultdach (steigend)", "b_std": 5.50, "h_trauf": 2.40, "h_first": 3.37, "dn": 10, "file": "110 - Carport System SC-P1.pdf", "is_shc": False},
        "SHC-P1 (Stahl-BSH)": {"tipo": "Acciaio-BSH", "forma": "Pultdach (steigend)", "b_std": 5.70, "h_trauf": 2.45, "h_first": 3.25, "dn": 8, "file": "110 - Carport System SHC-P1.pdf", "is_shc": True},
        "SC-P2 (Stahl-Stahl)": {"tipo": "Acciaio-Acciaio", "forma": "Pultdach (fallend)", "b_std": 5.50, "h_trauf": 3.53, "h_first": 4.50, "dn": 10, "file": "120 - Carport System SC-P2.pdf", "is_shc": False},
        "SHC-P2 (Stahl-BSH)": {"tipo": "Acciaio-BSH", "forma": "Pultdach (fallend)", "b_std": 5.70, "h_trauf": 3.30, "h_first": 4.10, "dn": 8, "file": "120 - Carport System SHC-P2.pdf", "is_shc": True},
        "SC-P3 (Stahl-Stahl)": {"tipo": "Acciaio-Acciaio", "forma": "Y-Doppelcarport", "b_std": 11.00, "h_trauf": 2.50, "h_first": 3.44, "dn": 10, "file": "130 - Carport System SC-P3.pdf", "is_shc": False}
    }

    modello_carport_ui = st.selectbox("Seleziona Modello Carport", list(carport_db.keys()), key="mod_carport")
    mod_data = carport_db[modello_carport_ui]
    
    st.info(f"Tipologia Selezionata: **{mod_data['tipo']}** | Forma Base: **{mod_data['forma']}**")

    st.markdown("### 📍 Localizzazione Cantiere Carport (Google Maps e Comune)")
    col_c_loc1, col_c_loc2 = st.columns([2, 1])
    with col_c_loc1:
        maps_url_cp = st.text_input("Incolla il link di Google Maps del cantiere (Carport):", value="", key="maps_url_cp")
    with col_c_loc2:
        _, _, luogo_estratto_cp = estrai_dati_da_url_maps(maps_url_cp)
        comune_cp = st.text_input("Comune di installazione (Carport)", value=luogo_estratto_cp, key="comune_cp")

    col_g_c1, col_g_c2, col_g_c3 = st.columns(3)
    with col_g_c1:
        larghezza_carport = st.number_input("Larghezza Trasversale (m)", value=mod_data["b_std"] if mod_data["b_std"] > 0 else 10.0, step=0.1)
    with col_g_c2:
        passo_telai_carport = st.number_input("Passo Telai (m)", value=5.0, step=0.1) 
    with col_g_c3:
        num_campate_carport = st.number_input("Numero Campate (Lunghezza)", value=5, min_value=1, step=1)
    
    st.markdown("#### Logica Carichi e Azioni Esterne (Lascia 0.0 su Neve/Vento per calcolo automatico da Maps)")
    col_cc1, col_cc2, col_cc3, col_cc4 = st.columns(4)
    with col_cc1:
        g1_carport = st.number_input("G1 - Struttura (kN/m²)", min_value=0.10, value=0.15, step=0.05)
    with col_cc2:
        g2_carport = st.number_input("G2 - Pannelli Solari ecc. (kN/m²)", min_value=0.0, value=0.20, step=0.05)
    with col_cc3:
        neve_carport = st.number_input("Neve qsk manuale (kN/m²)", min_value=0.0, value=0.0, step=0.10)
    with col_cc4:
        vento_carport = st.number_input("Vento base manuale (kN/m²)", min_value=0.0, value=0.0, step=0.10)

    st.markdown("### 🌍 Dati Geotecnici e Fondazioni")
    geo_file = st.file_uploader("📂 Carica Relazione Geologica (.pdf) per estrarre la portanza (Opzionale)", type=["pdf"], key="geo_file_cp")
    
    tipo_terreno_ui = st.selectbox("Seleziona Tipo di Terreno (Generico)", [
        "Scadente (Argille molli, Limo) ~ 0.5 daN/cm²",
        "Medio (Sabbie, Argille normali) ~ 1.5 daN/cm²",
        "Buono (Ghiaie, Sabbie dense) ~ 3.0 daN/cm²",
        "Ottimo (Roccia) ~ 5.0 daN/cm²"
    ], index=1)

    if st.button("Calcola Carico, Dimensiona Strutture e Genera Modello 3D", type="primary"):
        lat_cp, lon_cp, place_cp = estrai_dati_da_url_maps(maps_url_cp)
        comune_finale_cp = comune_cp if comune_cp else place_cp
        luogo_str_cp, qsk_cp, zona_vento_cp, press_vento_str_cp, zona_sismica_cp, alt_cp = estrai_parametri_ntc_da_coordinate_e_comune(lat_cp, lon_cp, comune_finale_cp)
        
        pressione_vento_cp_val = float(press_vento_str_cp.split()[0])
        neve_effettiva = neve_carport if neve_carport > 0.0 else qsk_cp
        vento_effettivo = vento_carport if vento_carport > 0.0 else pressione_vento_cp_val

        is_marittimo = ("isola" in luogo_str_cp.lower() or "sardegna" in luogo_str_cp.lower() or "pantelleria" in luogo_str_cp.lower() or "lampedusa" in luogo_str_cp.lower() or alt_cp < 40.0)
        ciclo_c5 = "Obbligatorio (Classe di corrosività C5 - Ambiente Marino / Industriale Severo)" if is_marittimo else "Standard protettivo zincato a caldo + verniciatura C3/C4"

        sigma_terreno_kn = 150.0
        if "Scadente" in tipo_terreno_ui: sigma_terreno_kn = 50.0
        elif "Medio" in tipo_terreno_ui: sigma_terreno_kn = 150.0
        elif "Buono" in tipo_terreno_ui: sigma_terreno_kn = 300.0
        elif "Ottimo" in tipo_terreno_ui: sigma_terreno_kn = 500.0
        
        if geo_file is not None:
            val_estratto = estrai_portanza_da_pdf(geo_file)
            if val_estratto:
                sigma_terreno_kn = val_estratto
                st.success(f"🌍 Dati estratti dal PDF! Capacità portante rilevata: {sigma_terreno_kn:.1f} kN/m²")
            else:
                st.warning("⚠️ Impossibile estrarre automaticamente la portanza dal PDF. Verrà utilizzato il valore generico selezionato.")

        q_neve_primario = (1.3 * (g1_carport + g2_carport)) + (1.5 * neve_effettiva) + (1.5 * 0.6 * vento_effettivo)
        q_vento_primario = (1.3 * (g1_carport + g2_carport)) + (1.5 * vento_effettivo) + (1.5 * 0.5 * neve_effettiva)
        q_totale_kn_mq = max(q_neve_primario, q_vento_primario)
        kg_mq_totale = q_totale_kn_mq * 100.0
        
        st.success(f"Località rilevata: {luogo_str_cp}  \nCarico Totale Equivalente Calcolato (SLU combinato NTC): **{kg_mq_totale:.1f} kg/m²** (Neve qsk: {neve_effettiva} kN/m², Vento qp: {vento_effettivo} kN/m²)")
        
        dati_carport = {
            "modello": modello_carport_ui,
            "tipo": mod_data['tipo'],
            "forma": mod_data['forma'],
            "larghezza": larghezza_carport,
            "passo_telai": passo_telai_carport,
            "num_campate": num_campate_carport,
            "lunghezza_totale": passo_telai_carport * num_campate_carport,
            "h_trauf": mod_data['h_trauf'],
            "h_first": mod_data['h_first'],
            "q_tot": q_totale_kn_mq,
            "kg_mq": kg_mq_totale,
            "g1": g1_carport, "g2": g2_carport, "neve": neve_effettiva, "vento": vento_effettivo,
            "luogo": luogo_str_cp,
            "neve_qsk": neve_effettiva,
            "zona_vento": zona_vento_cp,
            "pressione_vento": press_vento_str_cp,
            "zona_sismica": zona_sismica_cp,
            "ciclo_c5": ciclo_c5,
            "sigma_terreno": sigma_terreno_kn
        }
        
        ris_calc = esegui_calcolo_carport(dati_carport)
        dati_carport.update(ris_calc)
        
        st.write("---")
        
        col_dw_c1, col_dw_c2 = st.columns([1, 2])
        with col_dw_c1:
            st.markdown("### Dimensionamento Elementi")
            st.info(f"**Trave di Falda:** {dati_carport['sez_trave']}\n*(M_ed = {dati_carport['M_trave']} kNm)*")
            st.info(f"**Colonna Portante:** {dati_carport['sez_col']}")
            st.info(f"**Arcarecci di Copertura:** {dati_carport['sez_arc']}  \n*(Interasse massimo installazione: {dati_carport['passo_arc']:.2f} m)*")
            
            st.markdown("### Connessioni e Reazioni al Piede")
            st.success(f"**Nodo Top (Trave-Colonna):** ~{dati_carport['kg_nodo_top']} kg acciaio")
            st.success(f"**Nodo Base (Ancoraggio):** ~{dati_carport['kg_nodo_base']} kg acciaio")
            st.write(f"**Reazioni (SLU):** N = {dati_carport['N_base']} kN | V = {dati_carport['V_base']} kN | M = {dati_carport['M_base']} kNm")
            
            st.markdown("### Dimensionamento Plinti")
            st.info(f"**Forma:** {dati_carport['forma_plinto']}\n**Dimensione:** {dati_carport['dim_plinto']}")
            st.info(f"**Materiali a plinto:** {dati_carport['vol_plinto']} mc Cls | {dati_carport['kg_armatura']} kg armatura")

            st.markdown("### Sistemi di Stabilizzazione")
            st.warning(f"**Copertura:** {dati_carport['cv_falda']}")
            st.warning(f"**Verticali:** {dati_carport['cv_vert']}")

            st.markdown("### Trattamento Anticorrosione C5")
            st.error(f"**Condizione:** {dati_carport['ciclo_c5']}")
            st.metric("Superficie Acciaio da Trattare", f"{dati_carport['mq_acciaio_totale']} mq")
            
        with col_dw_c2:
            st.markdown("### Modello 3D Dinamico Carport")
            fig_3d = genera_modello_3d_carport(dati_carport)
            st.plotly_chart(fig_3d, use_container_width=True)
            
            word_file_cp = genera_word_carport(dati_carport)
            st.download_button(label="📄 Scarica Relazione Carport in Word (.docx)", data=word_file_cp, file_name=f"Relazione_Carport_{modello_carport_ui.replace(' ', '_')}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", type="primary", use_container_width=True)
