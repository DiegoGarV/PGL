#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Convierte UN log de métricas VR (.txt) en un Excel y dos CSV.

Uso (Windows):
  1) Ejecutar Abrir_convertidor.bat y seleccionar un .txt.
  2) O pasar el archivo en consola: py -3 Convertidor_Metricas_VR.py "Metrics_P001.txt"
No requiere instalar librerías externas: utiliza únicamente Python estándar.
"""

import csv
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
import re
import sys
from statistics import mean
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

TIME_RX = re.compile(r"^(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d\.\d+)\s*\|\s*(.+)$")
INTER_COLS = ["IdEvento", "Participante", "FechaHora", "Minijuego", "TipoAccion", "Resultado", "Dificultad",
              "TiempoEsperado_s", "TiempoReal_s", "DesfaseTiming_s", "ErrorTiming_s", "TiempoReaccion_s",
              "PrecisionEspacial", "PrecisionDireccion", "ManoCorrecta", "LineaOrigen"]
ADAPT_COLS = ["IdEvento", "FechaHora", "Minijuego", "DificultadAnterior", "DificultadNueva", "Motivo",
              "TasaExitoReciente", "TasaExito", "LineaOrigen"]
ESTADO_COLS = ["IdEvento", "FechaHora", "EstadoJugador", "LineaOrigen"]
ALL_COLS = ["IdEvento", "FechaHora", "Categoria", "Minijuego", "TipoAccion", "Resultado", "Dificultad",
            "EstadoJugador", "CambioDificultad", "Motivo", "ErrorTiming_s", "TiempoReaccion_s", "LineaOrigen", "TextoOriginal"]
FIELD_NUM = {"ExpectedTime": "TiempoEsperado_s", "ActualTime": "TiempoReal_s",
             "TimingOffset": "DesfaseTiming_s", "TimingError": "ErrorTiming_s",
             "ReactionTime": "TiempoReaccion_s", "SpatialAccuracy": "PrecisionEspacial",
             "DirectionAccuracy": "PrecisionDireccion"}
SUMMARY_FIELDS = {"Total Interactions": "Interacciones totales", "Successful": "Aciertos (Success)",
                  "Failed": "Fallos (Failed)", "Missed": "Omisiones (Missed)",
                  "Success Rate": "Tasa de éxito", "Average Timing Error": "Error timing medio (s)",
                  "Average Reaction Time": "Reacción media (s)", "Average Spatial Accuracy": "Precisión espacial media",
                  "Average Direction Accuracy": "Precisión dirección media"}


def parse_value(value):
    s = str(value).strip()
    if not s or s.upper() in ("N/A", "NULL", "NONE"):
        return None
    if s.lower() == "true":
        return True
    if s.lower() == "false":
        return False
    try:
        return float(s)
    except ValueError:
        return s


def parse_pairs(chunks):
    pairs = {}
    for chunk in chunks:
        if ":" in chunk:
            k, v = chunk.split(":", 1)
            pairs[k.strip()] = v.strip()
    return pairs


def read_log(path):
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    participant, started, summary, events = None, None, {}, []
    in_summary = False
    for lineno, raw in enumerate(text.splitlines(), start=1):
        line = raw.strip()
        if not line:
            continue
        if "SESSION SUMMARY" in line:
            in_summary = True
            continue
        if "END OF SESSION" in line:
            in_summary = False
            continue
        if line.startswith("===="):
            continue
        match = TIME_RX.match(line)
        if match:
            ts, body = match.groups()
            chunks = [c.strip() for c in body.split(" | ")]
            first = chunks[0]
            kv = parse_pairs(chunks)
            eid = len(events) + 1
            data = {"IdEvento": eid, "FechaHora": ts, "LineaOrigen": lineno, "TextoOriginal": line}
            if first.startswith("Minigame:"):
                data.update({"Categoria": "Interaction", "Minijuego": kv.get("Minigame", ""),
                             "TipoAccion": kv.get("Type", ""), "Resultado": kv.get("Outcome", ""),
                             "Dificultad": kv.get("Difficulty", ""), "Participante": participant or ""})
                for key, col in FIELD_NUM.items():
                    data[col] = parse_value(kv.get(key, "N/A"))
                data["ManoCorrecta"] = parse_value(kv.get("CorrectHand", "N/A"))
            elif first.startswith("DifficultyChanged:"):
                levels = kv.get("DifficultyChanged", "").split("->", 1)
                previous, new = ([x.strip() for x in levels] + [""])[:2]
                data.update({"Categoria": "DifficultyChanged", "DificultadAnterior": previous,
                             "DificultadNueva": new, "Motivo": kv.get("Reason", ""),
                             "Minijuego": kv.get("Minigame", ""),
                             "TasaExitoReciente": parse_value(kv.get("RecentSuccessRate", "N/A")),
                             "TasaExito": parse_value(kv.get("SuccessRate", "N/A"))})
            elif first.startswith("PlayerState:"):
                data.update({"Categoria": "PlayerState", "EstadoJugador": kv.get("PlayerState", "")})
            else:
                data.update({"Categoria": "Other", "Detalle": body})
            events.append(data)
            continue
        if ":" in line:
            key, value = (x.strip() for x in line.split(":", 1))
            if in_summary:
                summary[key] = parse_value(value)
            elif key == "Participant":
                participant = value
            elif key == "Session Start":
                started = value
    participant = participant or str(summary.get("Participant") or "SinID")
    for ev in events:
        if ev["Categoria"] == "Interaction":
            ev["Participante"] = participant
    return participant, started, summary, events


def only(events, category):
    return [event for event in events if event["Categoria"] == category]


def avg(values):
    valid = [float(v) for v in values if isinstance(v, (float, int)) and not isinstance(v, bool)]
    return mean(valid) if valid else None


def stat(group):
    total = len(group)
    counts = Counter(x.get("Resultado", "") for x in group)
    return {"total": total, "success": counts["Success"], "failed": counts["Failed"],
            "missed": counts["Missed"], "rate": counts["Success"] / total if total else None,
            "timing": avg(x.get("ErrorTiming_s") for x in group),
            "reaction": avg(x.get("TiempoReaccion_s") for x in group),
            "spatial": avg(x.get("PrecisionEspacial") for x in group),
            "direction": avg(x.get("PrecisionDireccion") for x in group)}


def dict_rows(rows, headers):
    return [[r.get(header) for header in headers] for r in rows]


def write_csv(path, headers, rows):
    with path.open("w", encoding="utf-8-sig", newline="") as fp:
        writer = csv.writer(fp)
        writer.writerow(headers)
        for row in rows:
            writer.writerow(["" if v is None else v for v in row])


# Construcción ligera de .xlsx con la biblioteca estándar de Python.
# Sin dependencias externas para facilitar su instalación y reutilización.
MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
ET.register_namespace("", MAIN)


def sub(parent, tag, attrs=None, text=None, ns=MAIN):
    elem = ET.SubElement(parent, f"{{{ns}}}{tag}", attrs or {})
    if text is not None:
        elem.text = str(text)
    return elem


def serialize(root):
    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    # Prefijos estándar en los archivos de relaciones del paquete OOXML.
    if root.tag.startswith(f"{{{CT}}}") or root.tag.startswith(f"{{{PKG}}}"):
        data = (data.replace(b"<ns0:", b"<")
                    .replace(b"</ns0:", b"</")
                    .replace(b"xmlns:ns0=", b"xmlns="))
    return data


def xlcol(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def formula(expr, cache):
    return ("FORMULA", expr.lstrip("="), cache)


def format_sheet(name, rows, header_at=1, title_at=None, pct_cols=(), dec_cols=(),
                 widths=None, freeze=None, filter_header=True):
    """rows son matrices regulares de valores, con formulas como tuplas especiales."""
    root = ET.Element(f"{{{MAIN}}}worksheet")
    maxcols = max([len(x) for x in rows] + [1])
    sub(root, "dimension", {"ref": f"A1:{xlcol(maxcols)}{max(1, len(rows))}"})
    views = sub(root, "sheetViews")
    view = sub(views, "sheetView", {"workbookViewId": "0"})
    if freeze is None:
        freeze = header_at
    sub(view, "pane", {"ySplit": str(freeze), "topLeftCell": f"A{freeze+1}", "activePane": "bottomLeft", "state": "frozen"})
    sub(root, "sheetFormatPr", {"defaultRowHeight": "17"})
    cols = sub(root, "cols")
    for i in range(1, maxcols + 1):
        width = widths[i - 1] if widths and i <= len(widths) else 19
        sub(cols, "col", {"min": str(i), "max": str(i), "width": str(width), "customWidth": "1"})
    data = sub(root, "sheetData")
    for rix, vals in enumerate(rows, 1):
        r = sub(data, "row", {"r": str(rix), "ht": "30" if rix == title_at else "22" if rix == header_at else "17", "customHeight": "1"})
        for cix, val in enumerate(vals, 1):
            if val is None:
                continue
            if val == "" and not (rix == header_at):
                continue
            pos = f"{xlcol(cix)}{rix}"
            style = 1 if rix == header_at else 2 if rix == title_at else (3 if cix in pct_cols else 4 if cix in dec_cols else 6 if rix % 2 else 0)
            attrs = {"r": pos, "s": str(style)}
            if isinstance(val, tuple) and len(val) == 3 and val[0] == "FORMULA":
                c = sub(r, "c", attrs)
                sub(c, "f", text=val[1])
                if val[2] is None:
                    c.set("t", "str")
                    sub(c, "v", text="")
                else:
                    sub(c, "v", text=float(val[2]) if isinstance(val[2], float) else val[2])
            elif isinstance(val, bool):
                c = sub(r, "c", dict(attrs, t="b"))
                sub(c, "v", text="1" if val else "0")
            elif isinstance(val, (int, float)):
                c = sub(r, "c", attrs)
                sub(c, "v", text=val)
            else:
                c = sub(r, "c", dict(attrs, t="inlineStr"))
                si = sub(c, "is")
                t = sub(si, "t", text=str(val))
                if str(val).startswith(" ") or str(val).endswith(" "):
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    if filter_header:
        sub(root, "autoFilter", {"ref": f"A{header_at}:{xlcol(maxcols)}{len(rows)}"})
    sub(root, "pageMargins", {"left": "0.4", "right": "0.4", "top": "0.7", "bottom": "0.7", "header": "0.3", "footer": "0.3"})
    return name, serialize(root)


def make_styles():
    root = ET.Element(f"{{{MAIN}}}styleSheet")
    fmts = sub(root, "numFmts", {"count": "2"})
    sub(fmts, "numFmt", {"numFmtId": "164", "formatCode": "0.00%"})
    sub(fmts, "numFmt", {"numFmtId": "165", "formatCode": "0.0000"})
    fonts = sub(root, "fonts", {"count": "3"})
    for textcolor, bold, size in [("233047", False, 10), ("FFFFFF", True, 10), ("152F4F", True, 15)]:
        ft = sub(fonts, "font")
        sub(ft, "sz", {"val": str(size)})
        sub(ft, "name", {"val": "Aptos"})
        if bold:
            sub(ft, "b")
        sub(ft, "color", {"rgb": "FF" + textcolor})
    fills = sub(root, "fills", {"count": "4"})
    sub(sub(fills, "fill"), "patternFill", {"patternType": "none"})
    sub(sub(fills, "fill"), "patternFill", {"patternType": "gray125"})
    x = sub(sub(fills, "fill"), "patternFill", {"patternType": "solid"})
    sub(x, "fgColor", {"rgb": "FF0F766E"})
    sub(x, "bgColor", {"indexed": "64"})
    x = sub(sub(fills, "fill"), "patternFill", {"patternType": "solid"})
    sub(x, "fgColor", {"rgb": "FFF0F7F7"})
    sub(x, "bgColor", {"indexed": "64"})
    borders = sub(root, "borders", {"count": "1"})
    b = sub(borders, "border")
    for side in ("left", "right", "top", "bottom", "diagonal"):
        sub(b, side)
    cs = sub(root, "cellStyleXfs", {"count": "1"})
    sub(cs, "xf", {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0"})
    xfs = sub(root, "cellXfs", {"count": "7"})
    specs = [(0,0,0), (1,2,0), (2,0,0), (0,0,164), (0,0,165), (1,2,0), (0,3,0)]
    for font, fill, numfmt in specs:
        xf = sub(xfs, "xf", {"numFmtId": str(numfmt), "fontId": str(font), "fillId": str(fill),
                                 "borderId": "0", "xfId": "0", "applyNumberFormat": "1" if numfmt else "0", "applyAlignment": "1"})
        sub(xf, "alignment", {"vertical": "center", "wrapText": "1" if fill == 2 else "0"})
    styles = sub(root, "cellStyles", {"count": "1"})
    sub(styles, "cellStyle", {"name": "Normal", "xfId": "0", "builtinId": "0"})
    return serialize(root)


def write_xlsx(path, sheets):
    workbook = ET.Element(f"{{{MAIN}}}workbook")
    tabs = sub(workbook, "sheets")
    rels = ET.Element(f"{{{PKG}}}Relationships")
    sub(rels, "Relationship", {"Id": "rId1", "Type": REL + "/styles", "Target": "styles.xml"}, ns=PKG)
    for i, (name, _) in enumerate(sheets, start=1):
        sub(tabs, "sheet", {"name": name, "sheetId": str(i), f"{{{REL}}}id": f"rId{i+1}"})
        sub(rels, "Relationship", {"Id": f"rId{i+1}", "Type": REL + "/worksheet", "Target": f"worksheets/sheet{i}.xml"}, ns=PKG)
    sub(workbook, "calcPr", {"calcId": "191029", "fullCalcOnLoad": "1"})
    package_rels = ET.Element(f"{{{PKG}}}Relationships")
    sub(package_rels, "Relationship", {"Id": "rId1", "Type": REL + "/officeDocument", "Target": "xl/workbook.xml"}, ns=PKG)
    cts = ET.Element(f"{{{CT}}}Types")
    sub(cts, "Default", {"Extension": "rels", "ContentType": "application/vnd.openxmlformats-package.relationships+xml"}, ns=CT)
    sub(cts, "Default", {"Extension": "xml", "ContentType": "application/xml"}, ns=CT)
    sub(cts, "Override", {"PartName": "/xl/workbook.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"}, ns=CT)
    sub(cts, "Override", {"PartName": "/xl/styles.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"}, ns=CT)
    for i in range(1, len(sheets) + 1):
        sub(cts, "Override", {"PartName": f"/xl/worksheets/sheet{i}.xml", "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"}, ns=CT)
    with ZipFile(path, "w", ZIP_DEFLATED, compresslevel=6) as z:
        z.writestr("[Content_Types].xml", serialize(cts))
        z.writestr("_rels/.rels", serialize(package_rels))
        z.writestr("xl/workbook.xml", serialize(workbook))
        z.writestr("xl/_rels/workbook.xml.rels", serialize(rels))
        z.writestr("xl/styles.xml", make_styles())
        for i, (_, data) in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", data)


def create_files(log_path):
    participant, started, summary, events = read_log(log_path)
    if not events:
        raise ValueError("No se encontraron eventos reconocibles. Verifica que es un log 'VR Metrics Log'.")
    interactions = only(events, "Interaction")
    dda = only(events, "DifficultyChanged")
    states = only(events, "PlayerState")
    others = only(events, "Other")
    stats = stat(interactions)

    # Tres salidas por sesión, sin mezclar participantes o archivos de origen.
    outdir = log_path.parent / "Convertidos"
    outdir.mkdir(exist_ok=True)
    stem = log_path.stem
    csv_inter = outdir / f"{stem}_interacciones.csv"
    csv_events = outdir / f"{stem}_todos_eventos.csv"
    xlsx_path = outdir / f"{stem}_analisis.xlsx"
    write_csv(csv_inter, INTER_COLS, dict_rows(interactions, INTER_COLS))
    all_rows = []
    for e in events:
        d = dict(e)
        if d["Categoria"] == "DifficultyChanged":
            d["CambioDificultad"] = f"{d.get('DificultadAnterior','')} -> {d.get('DificultadNueva','')}"
        all_rows.append(d)
    write_csv(csv_events, ALL_COLS, dict_rows(all_rows, ALL_COLS))

    # Hoja 1: totales reportados vs. recuento independiente de los eventos.
    r = [[f"Resumen de sesión VR · {participant}"], [],
         [f"Inicio: {started or 'No indicado'}     |     Fuente: {log_path.name}"], [],
         ["Indicador", "Reportado en log", "Calculado de eventos", None, "Otros registros", "Cantidad", "Descripción"]]
    last = len(interactions) + 1
    ir = f"'Interacciones'!$A$2:$A${last}"
    outcomes = f"'Interacciones'!$F$2:$F${last}"
    numeric = {"Interacciones totales": ("Total Interactions", formula(f"COUNTA({ir})", stats['total'])),
               "Aciertos (Success)": ("Successful", formula(f'COUNTIF({outcomes},"Success")', stats['success'])),
               "Fallos (Failed)": ("Failed", formula(f'COUNTIF({outcomes},"Failed")', stats['failed'])),
               "Omisiones (Missed)": ("Missed", formula(f'COUNTIF({outcomes},"Missed")', stats['missed'])),
               "Tasa de éxito": ("Success Rate", formula("IFERROR(C7/C6,0)", stats['rate'])),
               "Error timing medio (s)": ("Average Timing Error", formula(f'IFERROR(AVERAGE(\'Interacciones\'!$K$2:$K${last}),"")', stats['timing'])),
               "Reacción media (s)": ("Average Reaction Time", formula(f'IFERROR(AVERAGE(\'Interacciones\'!$L$2:$L${last}),"")', stats['reaction'])),
               "Precisión espacial media": ("Average Spatial Accuracy", formula(f'IFERROR(AVERAGE(\'Interacciones\'!$M$2:$M${last}),"")', stats['spatial'])),
               "Precisión dirección media": ("Average Direction Accuracy", formula(f'IFERROR(AVERAGE(\'Interacciones\'!$N$2:$N${last}),"")', stats['direction']))}
    for i, (label, (key, value)) in enumerate(numeric.items(), 6):
        extras = [("Cambios de dificultad", len(dda), "Eventos de adaptación de dificultad"),
                  ("Estados del jugador", len(states), "Cambios globales de estado"),
                  ("Eventos totales", len(events), "Incluye todos los tipos de eventos"),
                  ("Otros eventos", len(others), "Eventos sin categoría estándar")]
        extra = extras[i-6] if i-6 < len(extras) else (None, None, None)
        r.append([label, summary.get(key), value, None, *extra])
    r += [[], ["Observaciones"],
          ["N/A se conserva como celda vacía: no equivale a cero."],
          ["Si no hay valores de precisión espacial, el promedio recalculado queda vacío; el log puede reportar 0."],
          ["Los estados del jugador no se asignan a minijuegos cuando el log no los identifica."],
          ["Se cuentan eventos tal como aparecen, incluso si comparten marca temporal; no se eliminan duplicados."]]
    sheets = [format_sheet("Resumen", r, header_at=5, title_at=1, pct_cols=(),
                           widths=[38, 23, 26, 3, 25, 14, 45], freeze=5, filter_header=False)]

    # Hoja 2: rendimiento por minijuego.
    per_game = defaultdict(list)
    for e in interactions:
        per_game[e.get("Minijuego", "")].append(e)
    by_game = [["Rendimiento por minijuego"], [], [],
               ["Minijuego", "Interacciones", "Aciertos", "Fallos", "Omitidos", "TasaExito", "ErrorMedio_s",
                "ReaccionMedia_s", "CambiosDDA", "AccionesRegistradas"]]
    for i, (name, group) in enumerate(sorted(per_game.items()), 5):
        g = stat(group)
        mini = f"'Interacciones'!$D$2:$D${last}"
        val_error = f"'Interacciones'!$K$2:$K${last}"
        val_reac = f"'Interacciones'!$L$2:$L${last}"
        d_end = len(dda)+1
        dda_range = f"'Adaptaciones'!$C$2:$C${max(2,d_end)}"
        by_game.append([name,
            formula(f"COUNTIF({mini},$A{i})", g["total"]),
            formula(f'COUNTIFS({mini},$A{i},{outcomes},"Success")', g["success"]),
            formula(f'COUNTIFS({mini},$A{i},{outcomes},"Failed")', g["failed"]),
            formula(f'COUNTIFS({mini},$A{i},{outcomes},"Missed")', g["missed"]),
            formula(f'IFERROR(C{i}/B{i},0)', g['rate']),
            formula(f'IFERROR(AVERAGEIF({mini},$A{i},{val_error}),"")', g['timing']),
            formula(f'IFERROR(AVERAGEIF({mini},$A{i},{val_reac}),"")', g['reaction']),
            formula(f'COUNTIF({dda_range},$A{i})', sum(e.get('Minijuego') == name for e in dda)),
            ", ".join(sorted({e.get("TipoAccion", "") for e in group}))])
    sheets.append(format_sheet("Por minijuego", by_game, header_at=4, title_at=1, pct_cols=(6,),
                               dec_cols=(7,8), widths=[22,18,14,14,14,15,19,20,17,34]))

    # Hoja 3: mantener separadas las acciones de cada minijuego.
    by_pair = defaultdict(list)
    for e in interactions:
        by_pair[(e.get("Minijuego", ""), e.get("TipoAccion", ""))].append(e)
    per_action = [["Rendimiento por tipo de acción"], [], [],
                  ["Minijuego", "TipoAccion", "Interacciones", "Aciertos", "Fallos", "Omitidos", "TasaExito", "ErrorMedio_s", "ReaccionMedia_s"]]
    all_type = f"'Interacciones'!$E$2:$E${last}"
    mini = f"'Interacciones'!$D$2:$D${last}"
    err = f"'Interacciones'!$K$2:$K${last}"
    reac = f"'Interacciones'!$L$2:$L${last}"
    for i, ((name, act), group) in enumerate(sorted(by_pair.items()), 5):
        st = stat(group)
        filt = f"{mini},$A{i},{all_type},$B{i}"
        per_action.append([name, act,
            formula(f"COUNTIFS({filt})", st['total']),
            formula(f'COUNTIFS({filt},{outcomes},"Success")', st['success']),
            formula(f'COUNTIFS({filt},{outcomes},"Failed")', st['failed']),
            formula(f'COUNTIFS({filt},{outcomes},"Missed")', st['missed']),
            formula(f"IFERROR(D{i}/C{i},0)", st['rate']),
            formula(f'IFERROR(AVERAGEIFS({err},{filt}),"")', st['timing']),
            formula(f'IFERROR(AVERAGEIFS({reac},{filt}),"")', st['reaction'])])
    sheets.append(format_sheet("Por accion", per_action, header_at=4, title_at=1, pct_cols=(7,),
                               dec_cols=(8,9), widths=[23,23,19,14,14,14,15,20,22]))

    # Hojas 4-7: datos sin agregación, listos para filtros/gráficas.
    sheets.append(format_sheet("Interacciones", [INTER_COLS] + dict_rows(interactions, INTER_COLS),
                               dec_cols=(8,9,10,11,12,13,14),
                               widths=[12,17,27,20,21,14,14,21,19,21,20,20,20,22,17,17]))
    sheets.append(format_sheet("Adaptaciones", [ADAPT_COLS] + dict_rows(dda, ADAPT_COLS),
                               pct_cols=(7,8), widths=[12,27,20,21,19,31,21,18,16]))
    sheets.append(format_sheet("Estados jugador", [ESTADO_COLS] + dict_rows(states, ESTADO_COLS),
                               widths=[13,27,23,17]))
    sheets.append(format_sheet("Todos eventos", [ALL_COLS] + dict_rows(all_rows, ALL_COLS),
                               dec_cols=(11,12), widths=[13,26,20,20,21,16,17,18,23,30,20,20,17,90]))

    dictionary = [["Guía de lectura de los registros"], [], [], ["Campo", "Interpretación", "Precaución"],
                  ["IdEvento", "Orden de aparición en el archivo", "Permite relacionar interacciones, DDA y estados"],
                  ["FechaHora", "Fecha y hora del evento", "Hora local de la fuente; zona horaria no indicada"],
                  ["Participante", "Identificador de la sesión", "Conserva el ID de cada TXT para cruces con encuestas"],
                  ["Resultado", "Success, Failed o Missed", "Se conservan las categorías originales"],
                  ["TipoAccion", "Tipo concreto de interacción", "HammerTouch y HammerHit se mantienen separados"],
                  ["TiempoEsperado_s", "Tiempo objetivo declarado por el log", "Segundos según el registro"],
                  ["TiempoReal_s", "Tiempo de ejecución declarado por el log", "Puede estar ausente incluso con Success"],
                  ["DesfaseTiming_s", "Diferencia con signo entre tiempos", "Negativo: ejecución anticipada"],
                  ["ErrorTiming_s", "Error temporal reportado", "Promedio solo de valores disponibles"],
                  ["TiempoReaccion_s", "Tiempo de reacción reportado", "N/A no debe interpretarse como 0"],
                  ["PrecisionEspacial", "Precisión espacial si se registró", "Si todos los datos faltan, promedio vacío"],
                  ["PrecisionDireccion", "Precisión de dirección", "Solo tomar valores efectivamente reportados"],
                  ["ManoCorrecta", "True/False si la interacción tiene mano asociada", "Vacío = dato no disponible"],
                  ["Adaptaciones", "Cambios de dificultad", "Se mantienen motivo y tasa informada"],
                  ["Estados jugador", "Cambios globales del estado del jugador", "No inferir minijuego sin dato explícito"],
                  ["Resumen", "Reportado vs. calculado", "Sirve para validar coincidencia y detectar campos sin datos"],
                  ["LineaOrigen", "Línea del archivo .txt", "Permite comprobar eventos contra fuente original"],
                  ["TextoOriginal", "Evento tal como estaba en el .txt", "Útil ante eventos nuevos/desconocidos"]]
    sheets.append(format_sheet("Diccionario", dictionary, header_at=4, title_at=1,
                               widths=[26,54,64], filter_header=False))

    write_xlsx(xlsx_path, sheets)
    return xlsx_path, csv_inter, csv_events, stats, len(dda), len(states), len(events)


def select_log():
    """Menú de consola: funciona aunque Tkinter/ventanas estén bloqueadas."""
    if len(sys.argv) > 1:
        arg = sys.argv[1].strip().strip('"').strip("'")
        if arg.lower() in ('--help', '-h', '/?'):
            print('Uso: py -3 Convertidor_Metricas_VR_consola.py [ruta_al_log.txt]')
            return None
        return Path(arg).expanduser().resolve()

    folder = Path(__file__).resolve().parent
    options = []
    for candidate in sorted(folder.glob('*.txt')):
        # Solo mostrar registros VR, no archivos de instrucciones.
        try:
            with candidate.open('r', encoding='utf-8-sig', errors='replace') as reader:
                if 'VR Metrics Log' in reader.read(300):
                    options.append(candidate)
        except OSError:
            pass

    print('Carpeta del conversor:', folder)
    if options:
        print('\nArchivos de métricas encontrados:')
        for number, candidate in enumerate(options, 1):
            print(f'  {number}) {candidate.name}')
        print('\nEscribe el número del archivo o pega la ruta de otro .txt.')
    else:
        print('No se encontraron logs en esta carpeta.')
        print('Pega la ruta completa del archivo .txt que deseas convertir.')
    print('Ejemplo: C:\\Users\\Usuario\\Documents\\Metrics_P001.txt')
    try:
        answer = input('\nArchivo o número (Enter para cancelar): ').strip().strip('"').strip("'")
    except (EOFError, KeyboardInterrupt):
        print('\nOperación cancelada.')
        return None
    if not answer:
        print('No se seleccionó ningún archivo.')
        return None
    if answer.isdigit() and options:
        ix = int(answer) - 1
        if 0 <= ix < len(options):
            return options[ix]
        print(f'No existe la opción {answer}.')
        return None
    user_path = Path(answer).expanduser()
    if not user_path.is_absolute():
        user_path = folder / user_path
    return user_path.resolve()


def main():
    print('=========================================', flush=True)
    print('      CONVERTIDOR DE METRICAS VR', flush=True)
    print('=========================================', flush=True)
    selected = select_log()
    if selected is None:
        return 0
    if not selected.is_file() or selected.suffix.lower() != '.txt':
        print(f'ERROR: no se encontró un archivo .txt válido: {selected}', flush=True)
        return 1
    print(f'\nLeyendo: {selected.name}', flush=True)
    try:
        xlsx, ci, ce, stats, d, s, total = create_files(selected)
    except Exception as error:
        print('ERROR al convertir el archivo:', error, flush=True)
        return 1
    print('\nCONVERSION TERMINADA CORRECTAMENTE')
    print(f'  Interacciones: {stats["total"]}')
    print(f'  Cambios de dificultad: {d}')
    print(f'  Estados del jugador: {s}')
    print(f'  Eventos totales: {total}')
    print('\nArchivos guardados en:')
    print(f'  {xlsx.parent}')
    for file in (xlsx, ci, ce):
        print('  -', file.name)
    return 0


if __name__ == '__main__':
    sys.exit(main())
