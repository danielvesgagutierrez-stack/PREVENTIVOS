# -*- coding: utf-8 -*-
"""
Normaliza el cronograma de preventivos (XLSX) a un JSON que consume la PWA.

Uso:
    python tools/normalize_excel.py "ruta/al/CRONOGRAMA.xlsx" "data/equipos.json"
    python tools/normalize_excel.py CRONOGRAMA.xlsx salida.json --hoja=2026-B
Si no se indica --hoja y hay varias hojas con seriales, se pregunta cuál usar.

Reglas:
  - Los datos de EDIFICIO / PISO / OFICINA viven en celdas combinadas: se
    propagan el valor de la celda superior izquierda de cada rango combinado.
  - El serial es la clave única. Se guarda como TEXTO para no perder ceros.
  - Se separa el CATALOGO (inventario) de las JORNADAS (ejecucion).
  - El color/observacion del Excel se convierte en un ESTADO:
      azul -> ALTA, rojo -> AUSENTE, naranja/otras novedades -> NOVEDAD,
      verde o "sin novedades" -> OK, sin datos -> PENDIENTE.
  - Los seriales invalidos ("SS", vacios, duplicados) se reportan en "alertas".
"""
import json
import re
import sys
import unicodedata
import warnings
from datetime import date, datetime
from pathlib import Path

import openpyxl

warnings.filterwarnings("ignore", category=UserWarning)

AZUL = {"FF00B0F0"}
ROJO = {"FFFF0000", "FFFF3300"}
VERDE = {"FF92D050", "FF70AD47", "THEME:9"}
NARANJA = {"FFFFC000", "FFED7D31", "THEME:5", "THEME:7"}

MESES = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
    7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic",
}


def norm(texto):
    if texto is None:
        return ""
    return " ".join(str(texto).replace("\n", " ").split())


def clave(texto):
    t = norm(texto).upper()
    t = "".join(c for c in unicodedata.normalize("NFKD", t) if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]+", "-", t).strip("-")


def fecha(valor):
    if isinstance(valor, (datetime, date)):
        return valor.strftime("%Y-%m-%d")
    t = norm(valor)
    return t[:10] if t else ""


def color_de(celda):
    relleno = celda.fill
    if relleno is None or relleno.patternType is None:
        return ""
    rgb = relleno.fgColor.rgb
    if isinstance(rgb, str):
        return rgb.upper()
    if relleno.fgColor.type == "theme":
        return "THEME:" + str(relleno.fgColor.theme)
    return ""


def estado_de(celda, observacion):
    color = color_de(celda)
    obs = norm(observacion).lower()
    if color in AZUL:
        return "ALTA", observacion
    if color in ROJO or "no se encuentra" in obs or "no se encuentra" in obs.replace("ú", "u"):
        return "AUSENTE", observacion
    if color in NARANJA or color.startswith("THEME:5") or color.startswith("THEME:7"):
        return "NOVEDAD", observacion
    if "fuera de la universidad" in obs or "trabajando desde casa" in obs or "core 2 duo" in obs or "torre sin usar" in obs:
        return "NOVEDAD", observacion
    if color in VERDE:
        return "OK", "" if obs in ("sin novedades", "sin novedad") else norm(observacion)
    # Sin color reconocido NO es un registro: queda pendiente aunque tenga fecha o semestre.
    if obs and obs not in ("sin novedades", "sin novedad"):
        return "NOVEDAD", norm(observacion)
    return "PENDIENTE", ""


def elegir_hoja(wb, nombre_explicito=None):
    """Devuelve la hoja del mantenimiento. Si hay varias, pregunta (o lista)."""
    candidatas = []
    for ws in wb.worksheets:
        es_jornada = bool(re.match(r"^\d{4}-[AB]$", ws.title.strip()))
        tiene_serial = any(
            re.search("serial", str(c.value or ""), re.I)
            for fila in ws.iter_rows(min_row=1, max_row=min(12, ws.max_row))
            for c in fila
        )
        if es_jornada or tiene_serial:
            candidatas.append(ws.title)
    if nombre_explicito:
        if nombre_explicito not in wb.sheetnames:
            raise SystemExit(f"No existe la hoja '{nombre_explicito}'. Hojas: {wb.sheetnames}")
        return nombre_explicito
    if len(candidatas) <= 1:
        return candidatas[0] if candidatas else wb.sheetnames[0]
    print("El archivo tiene varias hojas con equipos:")
    for i, n in enumerate(candidatas, 1):
        print(f"  {i}. {n}")
    if sys.stdin.isatty():
        while True:
            opc = input("¿Cuál hoja es la del mantenimiento? [1]: ").strip()
            if not opc:
                opc = "1"
            if opc.isdigit() and 1 <= int(opc) <= len(candidatas):
                return candidatas[int(opc) - 1]
            print("Opción no válida")
    return candidatas[0]


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    hoja_arg = next((a.split("=", 1)[1] for a in sys.argv[1:] if a.startswith("--hoja=")), None)
    origen = Path(args[0] if len(args) > 0 else r"C:\Users\DAVEGU\Desktop\2-CRONOGRAMA 2026 DANIEL.xlsx")
    destino = Path(args[1] if len(args) > 1 else "data/equipos.json")
    wb = openpyxl.load_workbook(origen, data_only=True)
    hoja_elegida = elegir_hoja(wb, hoja_arg)
    print("Hoja seleccionada:", hoja_elegida)

    # Mapa de celdas combinadas -> valor de la esquina superior izquierda.
    combinadas = {}
    for ws in wb.worksheets:
        mapa = {}
        for rango in ws.merged_cells.ranges:
            valor = ws.cell(rango.min_row, rango.min_col).value
            if valor is None:
                continue
            for fila in range(rango.min_row, rango.max_row + 1):
                for col in range(rango.min_col, rango.max_col + 1):
                    mapa[(fila, col)] = valor
        combinadas[ws.title] = mapa

    catalogo = {}
    areas = {}
    jornadas = []
    alertas = []

    for ws in [wb[hoja_elegida]]:
        titulo = ws.title.strip()
        mapa = combinadas.get(titulo, {})
        match = re.match(r"^(\d{4})-([AB])$", titulo)
        anio = match.group(1) if match else (re.search(r"\d{4}", titulo) or [str(datetime.now().year)])[0]
        semestre = match.group(2) if match else ""
        items = []
        vistos = {}

        for fila in range(2, ws.max_row + 1):
            def valor(col):
                celda = ws.cell(fila, col)
                v = celda.value
                if v is None:
                    v = mapa.get((fila, col))
                return norm(v)

            edificio, piso, oficina = valor(1), valor(2), valor(3)
            serial = valor(4)
            tipo = valor(5)
            if not serial and not tipo and not oficina:
                continue

            es_valido = bool(re.fullmatch(r"\d{9,16}", serial))
            if not es_valido:
                # Sin serial no hay equipo: se reporta pero NO cuenta en estadísticas.
                alertas.append({
                    "jornada": titulo, "fila": fila, "serial": serial,
                    "motivo": "Serial no numerico o vacio (posible 'sin serial')",
                })
                continue

            area_id = clave("|".join([edificio, piso, oficina])) or "sin-area"
            if area_id not in areas:
                areas[area_id] = {
                    "id": area_id,
                    "edificio": edificio,
                    "piso": piso,
                    "oficina": oficina,
                }

            observacion = valor(12)
            if not any(ws.cell(fila, c).value for c in (7, 9, 10, 11, 12)) and color_de(ws.cell(fila, 4)) == "":
                estado, obs = "PENDIENTE", ""
            else:
                estado, obs = estado_de(ws.cell(fila, 4), observacion)

            programada = fecha(ws.cell(fila, 6).value)
            ejecutada = fecha(ws.cell(fila, 7).value)
            tecnico = valor(8)
            duracion = valor(10)

            item = {
                "serial": serial,
                "areaId": area_id,
                "edificio": edificio,
                "piso": piso,
                "oficina": oficina,
                "tipo": tipo,
                "estado": estado,
                "novedad": obs,
                "programada": programada,
                "ejecutada": ejecutada,
                "tecnico": tecnico,
                "duracion": duracion,
                "filaExcel": fila,
                "serialValido": es_valido,
            }
            items.append(item)

            if serial in vistos:
                alertas.append({
                    "jornada": titulo, "fila": fila, "serial": serial,
                    "motivo": f"Serial duplicado (primera vez en fila {vistos[serial]})",
                })
            else:
                vistos[serial] = fila

            if serial not in catalogo:
                catalogo[serial] = {
                    "serial": serial,
                    "areaId": area_id,
                    "edificio": edificio,
                    "piso": piso,
                    "oficina": oficina,
                    "tipo": tipo,
                    "serialValido": es_valido,
                }

        tecnico = next((i["tecnico"] for i in items if i["tecnico"]), "")
        fechas = [i["programada"] for i in items if i["programada"]]
        jornadas.append({
            "id": clave(titulo),
            "nombre": f"Preventivo {semestre} {anio}",
            "semestre": semestre,
            "anio": anio,
            "tecnico": tecnico,
            "desde": min(fechas) if fechas else "",
            "hasta": max(fechas) if fechas else "",
            "items": items,
        })

    total_items = sum(len(j["items"]) for j in jornadas)
    salida = {
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "fuente": origen.name,
        "hoja": hoja_elegida,
        "columnas": {
            "edificio": 0, "piso": 1, "oficina": 2, "serial": 3, "tipo": 4,
            "programada": 5, "ejecutada": 6, "tecnico": 7, "semestre": 8,
            "duracion": 9, "seguimiento": 10, "observaciones": 11, "headerRow": 0,
        },
        "areas": sorted(areas.values(), key=lambda a: (a["edificio"], a["piso"], a["oficina"])),
        "catalogo": [catalogo[k] for k in sorted(catalogo)],
        "jornadas": jornadas,
        "alertas": alertas,
        "resumen": {
            "equiposUnicos": len(catalogo),
            "areas": len(areas),
            "jornadas": len(jornadas),
            "registros": total_items,
            "alertas": len(alertas),
        },
    }
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_text(json.dumps(salida, ensure_ascii=False, indent=1), encoding="utf-8")
    print("OK ->", destino)
    print(json.dumps(salida["resumen"], ensure_ascii=False, indent=2))
    for a in alertas[:12]:
        print("  alerta:", a["jornada"], "fila", a["fila"], a["serial"], "-", a["motivo"])
    if len(alertas) > 12:
        print("  ... y", len(alertas) - 12, "mas")


if __name__ == "__main__":
    main()
