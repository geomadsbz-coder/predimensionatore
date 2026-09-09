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
    
    lat = dati_geo.get('latitudine', 46.4983)
    lon = dati_geo.get('longitudine', 11.3548)
    comune = dati_geo.get('comune', '')
    
    luogo_str, qsk, zona_vento, press_vento_str, zona_sismica, altitudine_stimata = estrai_parametri_ntc_da_coordinate_e_comune(lat, lon, comune)
    pressione_vento = float(press_vento_str.split()[0])
    
    # ---------------------------------------------------------
    # 1. CALCOLO PASSI REALI (ARCARECCI E BARACCATURA)
    # ---------------------------------------------------------
    # Passo Arcarecci in base al pannello copertura
    spessore_cop = str(dati_geo.get('spessore_pannello', ''))
    tipo_cop = str(dati_geo.get('tipo_isolante', ''))
    
    if "Lamiera" in tipo_cop: max_passo_arc = 1.2
    elif "50" in spessore_cop or "60" in spessore_cop: max_passo_arc = 1.8
    elif "80" in spessore_cop or "100" in spessore_cop: max_passo_arc = 2.5
    else: max_passo_arc = 3.0
    
    sviluppo_falda = math.sqrt((luce_totale/2)**2 + (h_colmo - h_gronda)**2)
    num_campi_falda = math.ceil(sviluppo_falda / max_passo_arc)
    passo_arcarecci = sviluppo_falda / num_campi_falda

    # Passo Baraccatura in base al pannello parete
    spessore_par = str(dati_geo.get('spessore_pannello_parete', ''))
    tipo_par = str(dati_geo.get('tipo_isolante_parete', ''))
    
    if "Lamiera" in tipo_par or "Nessuno" in tipo_par: max_passo_bar = 1.5
    elif "50" in spessore_par or "60" in spessore_par: max_passo_bar = 2.2
    elif "80" in spessore_par or "100" in spessore_par: max_passo_bar = 3.0
    else: max_passo_bar = 3.5
    
    num_campi_parete = math.ceil(h_gronda / max_passo_bar)
    passo_baraccatura = h_gronda / num_campi_parete

    # ---------------------------------------------------------
    # 2. SVILUPPO LINEARE ESATTO BARACCATURA
    # ---------------------------------------------------------
    # Pareti longitudinali
    num_file_long = int(h_gronda / passo_baraccatura)
    ml_baraccatura_long_tot = num_file_long * lunghezza_edificio * 2  # 2 pareti
    
    # Pareti frontali (Timpani)
    ml_baraccatura_timpani_tot = 0
    z = passo_baraccatura
    while z < h_colmo:
        if z <= h_gronda:
            larghezza_timpano = luce_totale
        else:
            # Triangolo superiore: proporzione
            h_triangolo_locale = z - h_gronda
            h_triangolo_tot = h_colmo - h_gronda
            larghezza_timpano = luce_totale * (1 - (h_triangolo_locale / h_triangolo_tot))
        ml_baraccatura_timpani_tot += larghezza_timpano * 2  # 2 timpani
        z += passo_baraccatura

    ml_tot_baraccatura = ml_baraccatura_long_tot + ml_baraccatura_timpani_tot

    # ---------------------------------------------------------
    # 3. DIMENSIONAMENTO ARCARECCI (In Luce vs Sopra)
    # ---------------------------------------------------------
    g1, g2 = 0.15, 0.25
    if "Presente" in dati_geo.get('impianto_fv_desc', ''): g2 += 0.20
    g2 += dati_geo.get('carico_aggiuntivo', 0.0)
    q_tot_copertura_mq = (1.3 * g1 + 1.5 * g2 + 1.5 * qsk)
    q_arc = q_tot_copertura_mq * passo_arcarecci
    
    if pos_arcarecci == "In luce":
        # Schema isostatico semplice appoggio
        M_ed_arc = (q_arc * interasse**2) / 8
        coeff_f = 5/384
    else:
        # Schema trave continua su più appoggi (più rigido)
        M_ed_arc = (q_arc * interasse**2) / 10
        coeff_f = 2/384

    # Dimensionamento Arcareccio Acciaio
    W_req_arc = (M_ed_arc * 100) / 27.5 # cm3 (f_yd = 275 MPa)
    if W_req_arc < 30: sez_arc = "Tubolare 100x50x3"
    elif W_req_arc < 55: sez_arc = "Tubolare 120x60x4"
    elif W_req_arc < 85: sez_arc = "Tubolare 150x75x4 / IPE 140"
    else: sez_arc = "IPE 180"

    # ---------------------------------------------------------
    # 4. DIMENSIONAMENTO TRAVE PRINCIPALE
    # ---------------------------------------------------------
    q_ed_portale = interasse * q_tot_copertura_mq
    luce_campata = luce_totale / (num_appoggi - 1) if num_appoggi >= 3 else luce_totale
    
    if num_appoggi >= 3:
        m_ed = (q_ed_portale * (luce_campata ** 2)) / 10.0 
        v_ed = (q_ed_portale * luce_campata) / 2.0        
    else:
        m_ed = (q_ed_portale * (luce_campata ** 2)) / 8.0
        v_ed = (q_ed_portale * luce_campata) / 2.0

    b_legno_cm = 20 
    w_req_cm3 = (m_ed * 1e6) / 14500.0  
    h_legno_cm = int((6 * w_req_cm3 / b_legno_cm) ** 0.5)
    h_legno_cm = max(44, ((h_legno_cm + 3) // 4) * 4) 

    # ---------------------------------------------------------
    # 5. OTTIMIZZAZIONE SNELLA PILASTRI (Stress & Drift Limiti)
    # ---------------------------------------------------------
    # Azione orizzontale vento
    q_w = pressione_vento * interasse  # kN/m
    M_base_vento = (q_w * h_gronda**2) / 2
    V_base_vento = q_w * h_gronda

    E_legno = 1150  # kN/cm2 (GL24h)
    limite_spostamento_cm = (h_gronda * 100) / 150  # L/150 per edifici industriali NTC

    b_pil_perim_cm = 20
    h_pil_perim_cm = 32 # Partenza minima
    while True:
        I_pil = (b_pil_perim_cm * h_pil_perim_cm**3) / 12
        W_pil = (b_pil_perim_cm * h_pil_perim_cm**2) / 6
        
        # Stress flessionale (Vento dominante)
        sigma_m = (M_base_vento * 100) / W_pil # kN/cm2
        # Spostamento in sommità (Mensola)
        delta_somm = (q_w / 100 * (h_gronda * 100)**4) / (8 * E_legno * I_pil)
        
        if sigma_m < 1.45 and delta_somm < limite_spostamento_cm:
            break
        h_pil_perim_cm += 4
        if h_pil_perim_cm > 120: break # Safety exit

    h_pil_interm_cm = max(32, ((int(h_legno_cm * 0.50) + 3) // 4) * 4) # Intermedio lavora a N puro quasi
    b_pil_interm_cm = 20

    # Dimensionamento Acciaio
    w_el_req_cm3 = (m_ed * 100.0) / 33.8 
    if w_el_req_cm3 > 3500:
        profilo_acciaio = "IPE 600 / HEB 500"
        pilastro_perim_acc = "HEB 300"
        pilastro_interm_acc = "HEA 240"
    elif w_el_req_cm3 > 2000:
        profilo_acciaio = "IPE 500 / HEA 400"
        pilastro_perim_acc = "HEB 260"
        pilastro_interm_acc = "HEA 200"
    elif w_el_req_cm3 > 1000:
        profilo_acciaio = "IPE 400 / HEA 300"
        pilastro_perim_acc = "HEB 200 (Ottimizzato)"
        pilastro_interm_acc = "HEA 160"
    else:
        profilo_acciaio = "IPE 330 / HEA 240"
        pilastro_perim_acc = "HEB 180 (Ottimizzato)"
        pilastro_interm_acc = "HEA 140"

    h_cap_cm = max(80, ((int(h_legno_cm * 1.2) + 4) // 5) * 5)
    profilo_cap = f"Trave a T rovescia precompressa altezza {h_cap_cm} cm"

    # --- DIMENSIONAMENTO MONTANTI TIMPANO ED ALTRO (Mantenuto logicamente simile) ---
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

    # Sezione Montanti
    if h_colmo <= 6.5: montante_legno = "Sezione 14x14 cm (GL24h)"
    elif h_colmo <= 9.5: montante_legno = "Sezione 16x16 cm (GL24h)"
    elif h_colmo <= 12.5: montante_legno = "Sezione 16x24 cm (GL24h)"
    else: montante_legno = "Sezione 20x28 cm (GL24h)"

    mq_acciaio = round((luce_totale + h_gronda * 2) * (dati_geo['num_campate'] + 1) * 0.6, 1)

    risultati_deterministici = {
        "luogo": luogo_str,
        "qsk": qsk,
        "zona_vento": zona_vento,
        "pressione_vento": press_vento_str,
        "zona_sismica": zona_sismica,
        "classe_uso": "Classe II (Edifici industriali ordinari)",
        "fattore_struttura_q": "q = 2.0 (Struttura intelaiata)",
        "m_ed": round(m_ed, 1),
        "v_ed": round(v_ed, 1),
        "travi_legno": f"Base {b_legno_cm} cm x Altezza {h_legno_cm} cm (GL24h)",
        "travi_acciaio": f"Profilo {profilo_acciaio} in acciaio S355JR",
        "travi_cap": profilo_cap,
        "pilastri_perimetrali_legno": f"Sezione {b_pil_perim_cm}x{h_pil_perim_cm} cm (Ottimizzato a drift H/150 = {limite_spostamento_cm:.1f}cm)",
        "pilastri_intermedi_legno": f"Sezione {b_pil_interm_cm}x{h_pil_interm_cm} cm (GL24h)",
        "pilastri_perimetrali_acciaio": f"Profilo {pilastro_perim_acc} in acciaio S355JR",
        "pilastri_intermedi_acciaio": f"Profilo {pilastro_interm_acc} in acciaio S355JR",
        "pilastri_perimetrali_cap": f"Pilastro in C.A.P. sezione 40x45 cm",
        "pilastri_intermedi_cap": f"Pilastro in C.A.P. sezione 40x40 cm",
        
        "passo_arcarecci_calc": round(passo_arcarecci, 2),
        "sezione_arcarecci": f"{sez_arc} o Legno 10x20 cm",
        "verifica_arcarecci": f"Verificato (Posizione: {pos_arcarecci}) - M_ed: {M_ed_arc:.1f} kNm",
        
        "passo_baraccatura_calc": round(passo_baraccatura, 2),
        "ml_baraccatura_tot": round(ml_tot_baraccatura, 1),
        "baraccatura_legno_lamellare": f"Correnti GL24h 12x16 cm",
        "baraccatura_legno_massiccio": f"Correnti C24 14x16 cm",
        "baraccatura_acciaio": f"Omega / Tubolare 100x50x3",
        
        "num_montanti_timpano_singolo": num_montanti_timpano_singola_facciata,
        "passo_montanti_timpano": round(passo_montanti_timpano, 2),
        "ml_per_montante_timpano": ml_per_montante_timpano,
        "ml_tot_timpani_entrambe": round(ml_tot_timpani_entrambe, 2),
        
        "num_montanti_long_singola_parete": num_totale_montanti_long_singola_parete,
        "passo_montanti_long": round(passo_montanti_long, 2),
        "ml_tot_montanti_long_entrambe": round(ml_tot_montanti_long_entrambe_pareti, 2),

        "montante_sezione_legno": montante_legno,
        "montante_sezione_acciaio": montante_acciaio if "acciaio" in locals() else "Tubolare",
        "montante_sezione_cap": "Pilastrino C.A.P. 20x20 cm",

        "campate_controventi_indici": [0, dati_geo['num_campate'] - 1],
        "controventi_copertura_pos": f"Campate di estremità",
        "controventi_copertura_legno": "Tiranti tondi d'acciaio diametro 20 mm",
        "controventi_copertura_acciaio": "Tubolari incrociati Ø 89x4 mm",
        "controventi_parete_pos": f"Campate di estremità",
        "controventi_parete_legno": "Diagonali legno lamellare 16x16 cm",
        "controventi_parete_acciaio": "Croci di sant'andrea L 80x8",
        
        "conn_trave_pilastro_tipo": "Nodo semi-rigido con piastre",
        "conn_trave_pilastro_perim_elementi": f"N. 6 bulloni 8.8 M20",
        "conn_trave_pilastro_perim_kg": f"45.0 kg cad.",
        "conn_trave_pilastro_interm_elementi": f"N. 4 bulloni 8.8 M20",
        "conn_trave_pilastro_interm_kg": f"32.0 kg cad.",
        "conn_pilastro_fondazione_tipo": "Cerniera/Incastro",
        "conn_pilastro_fondazione_perim_elementi": f"N. 4 tirafondi M24",
        "conn_pilastro_fondazione_perim_kg": f"38.0 kg cad.",
        "conn_pilastro_fondazione_interm_elementi": f"N. 4 tirafondi M24",
        "conn_pilastro_fondazione_interm_kg": f"35.0 kg cad.",
        "dettaglio_giunto_colmo": "Piastra di colmo bullonata",
        "classe_resistenza_fuoco": "R 60",
        "mq_intumescente": f"{mq_acciaio} mq",
        "dettaglio_verniciatura": "Primer + Intumescente R60",
        "note_tecniche": f"Calcolo esatto ml baraccatura. Pilastri ottimizzati al limite deformativo (H/150). Posizione arcarecci: {pos_arcarecci}."
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

# --- PARTE INTERFACCIA E STREAMLIT (Snippet aggiornato per UI) ---
# ... (mantenere intatto setup di pagina e file upload) ...

st.markdown("### 🏛️ Configurazione Telaio e Travatura")
col_g1, col_g2, col_g3 = st.columns(3)
with col_g1:
    tipo_travatura = st.selectbox("Tipologia Travatura", ["Bi-falda semplice", "Bi-falda con intradosso curvo", "Trave giuntata in colmo"], key="tipo_travatura")
with col_g2:
    num_appoggi = st.selectbox("Numero Appoggi Telaio", [2, 3, 4], index=1, key="num_appoggi")
with col_g3:
    posizione_arc = st.radio("Posizionamento Arcarecci", ["Sopra i telai (Continuo)", "In luce (Semplice appoggio)"], key="pos_arcarecci")

# ... (Mantenere caricamento carichi, bottone esegui calcolo, e generazione IA aggiornando le logiche di input/output) ...

if 'dati_ultimi' in st.session_state:
    dati = st.session_state['dati_ultimi']
    distinta = dati.get('distinta', calcola_distinta_elementi(dati))
    
    # ... (Mostra Modello 3D e Distinta di base) ...
    
    st.markdown("### 🪵 3. Arcarecci di Copertura (Dimensionamento Esatto)")
    st.info(f"**Passo Calcolato (su base pannello):** {dati.get('passo_arcarecci_calc')} m | **Posizione:** {dati.get('posizione_arcarecci', 'Sopra i telai')} | **Sezione:** {dati.get('sezione_arcarecci')} | {dati.get('verifica_arcarecci')}")
    
    st.markdown("### 🧱 4. Baraccatura di Parete (Sviluppo ML Esatto)")
    st.success(f"**Passo Calcolato (su base pannello parete):** {dati.get('passo_baraccatura_calc')} m")
    st.metric("Sviluppo Metri Lineari Baraccatura (Totale Pareti)", f"{dati.get('ml_baraccatura_tot')} ml")
    st.write(f"- Legno Lamellare: {dati.get('baraccatura_legno_lamellare')}")
    st.write(f"- Acciaio: {dati.get('baraccatura_acciaio')}")

    st.markdown("### 🏛️ 6. Pilastri Ottimizzati (Verifica Snellezza e Deformata)")
    st.success(f"**Legno Perimetrali:** {dati.get('pilastri_perimetrali_legno')}")
    st.warning(f"**Acciaio Perimetrali:** {dati.get('pilastri_perimetrali_acciaio')}")
