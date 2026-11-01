#!/usr/bin/env python3
"""
Convierte el archivo Excel "Tarjeta Clara card data" (pestaña "Card Offers")
en los dos archivos que usa el sitio:

  - datos-tarjetas.json  -> versión legible, con indentación, SOLO para
                             nosotros (revisar/depurar). El sitio NO carga
                             este archivo.
  - datos-tarjetas.js    -> versión comprimida (sin espacios) que sí carga
                             el sitio. Trae únicamente los campos que el
                             código de index.html realmente usa -- menos
                             datos que descargar en el celular.

Se vuelve a correr cada vez que haya una versión nueva del Excel (por
ejemplo, cuando el CFPB publique su próxima encuesta semestral).

Uso: python3 generar-datos.py ruta/al/archivo.xlsx
"""
import openpyxl
import json
import re
import math
import sys

def clean(v):
    if v is None:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, str):
        v = v.strip()
        return v if v else None
    return v

def yesno(v):
    v = clean(v)
    return str(v).strip().lower() == 'yes' if v else False

def num(v):
    v = clean(v)
    if v is None:
        return None
    try:
        return round(float(v), 4)
    except (TypeError, ValueError):
        return None

def pct(v):
    n = num(v)
    return round(n * 100, 2) if n is not None else None

def slugify(text):
    text = text.lower()
    text = re.sub(r'[^a-z0-9]+', '-', text)
    return text.strip('-')

def main(ruta_excel):
    wb = openpyxl.load_workbook(ruta_excel, data_only=True)
    ws = wb['Card Offers']

    tarjetas = []
    seen = {}
    for row in ws.iter_rows(min_row=2, max_row=348, values_only=True):
        (issuer, card_name, situation_tags, apr_low, apr_high, intro_apr_purch, intro_apr_bt,
         intro_period_months, bt_fee_pct, annual_fee, fx_fee_pct, min_credit_tier, rewards_type,
         rewards_notes, source_url, date_verified, bt_fee_min, rewards_detail, features,
         availability, states, secured, special_elig, elig_notes, card_website,
         apr_no_score, apr_poor, apr_good, apr_great, intro_period_purchases) = row

        issuer = clean(issuer)
        card_name = clean(card_name)
        if not issuer or not card_name:
            continue

        base_id = slugify(f"{issuer}-{card_name}")[:80]
        n = seen.get(base_id, 0)
        seen[base_id] = n + 1
        tid = base_id if n == 0 else f"{base_id}-{n+1}"

        tags = [t.strip() for t in (clean(situation_tags) or '').split(',') if t.strip()]

        # Solo los campos que index.html realmente usa (ver obtenerRecomendaciones,
        # calificaPorCredito, calificaPorObjetivo, puntuarTarjeta, obtenerExplicacion,
        # mostrarResultado y calcularAhorroTransferencia). Si agregas una función
        # nueva que necesite otro campo del Excel, agrégalo aquí también.
        tarjeta = {
            "id": tid,
            "emisor": issuer,
            "producto": card_name,
            "etiquetas_situacion": tags,
            "apr_compra": {"baja": pct(apr_low), "alta": pct(apr_high)},
            "apr_intro_transferencia": pct(intro_apr_bt),
            "periodo_intro_transferencia_meses": num(intro_period_months),
            "cuota_transferencia_pct": pct(bt_fee_pct),
            "cuota_transferencia_minima": num(bt_fee_min),
            "cuota_anual": num(annual_fee) or 0,
            "cargo_transaccion_extranjera_pct": pct(fx_fee_pct) or 0,
            "nivel_credito_minimo_texto": clean(min_credit_tier),
            "tipo_recompensas": clean(rewards_type),
            "tarjeta_segura": yesno(secured),
            "requiere_elegibilidad_especial": yesno(special_elig),
            "sitio_web": clean(card_website) or clean(source_url),
        }
        tarjetas.append(tarjeta)

    salida = {
        "fuente": "CFPB TCCP Survey, vía hoja de cálculo 'Tarjeta Clara card data' (194 emisores curados, enfocados en TX/CA/FL)",
        "fecha_archivo": "2025-12-31",
        "total_tarjetas": len(tarjetas),
        "tarjetas": tarjetas,
    }

    # Versión legible, solo para nosotros.
    with open('datos-tarjetas.json', 'w', encoding='utf-8') as f:
        json.dump(salida, f, ensure_ascii=False, indent=2)

    # Versión comprimida: la que de verdad carga el sitio.
    with open('datos-tarjetas.js', 'w', encoding='utf-8') as f:
        f.write('// GENERADO AUTOMÁTICAMENTE por generar-datos.py -- no lo edites a mano.\n')
        f.write('// Para actualizarlo: python3 generar-datos.py "Tarjeta Clara card data.xlsx"\n')
        f.write('const datosTarjetas=')
        json.dump(salida, f, ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')

    print(f"Listo: {len(tarjetas)} tarjetas procesadas.")

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Uso: python3 generar-datos.py ruta/al/archivo.xlsx")
        sys.exit(1)
    main(sys.argv[1])
