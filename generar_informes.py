#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador de informes + portal web
Concentracion Educativa del Sur de Montelibano
Examen Final de Primer Periodo 2026
"""

import pandas as pd
import openpyxl
import os, json, re, unicodedata
from pathlib import Path

# ─── Config ───────────────────────────────────────────────────────────────────
BASE_DIR   = Path("/home/froylan/Documents/Pruebas finales 2026")
PORTAL_DIR = BASE_DIR / "portal"

COLEGIO = "Concentracion Educativa del Sur de Montelibano"
EXAMEN  = "Examen Final de Primer Periodo"
ANIO    = "2026"

AREAS_PREGUNTAS = {
    "Matematicas":       list(range(1,  21)),
    "Etica":             list(range(21, 28)),
    "Lengua Castellana": list(range(28, 48)),
    "Ed. Fisica":        list(range(48, 55)),
    "Ciencias Sociales": list(range(55, 75)),
    "Gestion":           list(range(75, 82)),
    "Ciencias Naturales":list(range(82, 102)),
    "Artistica":         list(range(102,109)),
    "Ingles":            list(range(109,129)),
    "Religion":          list(range(129,136)),
    "Tecnologia":        list(range(136,143)),
}

AREAS_PREGUNTAS_DECIMO = {
    "Matematicas":       list(range(1,  21)),
    "Filosofia":         list(range(21, 27)),
    "Lengua Castellana": list(range(27, 47)),
    "Etica":             list(range(47, 53)),
    "Ciencias Sociales": list(range(53, 73)),
    "C. Politicas":      list(range(73, 79)),
    "Quimica":           list(range(79, 89)),
    "Ed. Fisica":        list(range(89, 95)),
    "Fisica":            list(range(95, 105)),
    # Artistica (Q105-Q110): no hubo prueba en Decimo — omitida
    "Religion":          list(range(111, 117)),
    "Ingles":            list(range(117, 137)),
    "Tecnologia":        list(range(137, 143)),
}

COLORES_AREA = {
    "Matematicas":        "#4472C4",
    "Etica":              "#ED7D31",
    "Lengua Castellana":  "#538135",
    "Ed. Fisica":         "#C00000",
    "Ciencias Sociales":  "#7030A0",
    "Gestion":            "#00B050",
    "Ciencias Naturales": "#0070C0",
    "Artistica":          "#D63384",
    "Ingles":             "#D4A017",
    "Religion":           "#833C00",
    "Tecnologia":         "#404040",
    "Filosofia":         "#9B59B6",
    "C. Politicas":      "#1ABC9C",
    "Quimica":           "#E74C3C",
    "Fisica":            "#F39C12",
}

GRADE_NAMES = {
    "6": "Sexto",   "7": "Septimo",
    "8": "Octavo",  "9": "Noveno", "10": "Decimo",
}

# ─── Helpers ──────────────────────────────────────────────────────────────────

def sanitize(s):
    s = ''.join(c for c in unicodedata.normalize('NFD', str(s))
                if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^A-Za-z0-9_\-]', '_', s)

def grupo_desde_id(sid):
    sid = str(sid).strip()
    if sid.startswith('10'):
        gd = '10'; gr = sid[2] if len(sid) > 2 else '?'
    else:
        gd = sid[0]; gr = sid[2] if len(sid) > 2 else '?'
    gn = GRADE_NAMES.get(gd, f"{gd}o")
    gl = chr(ord('A') + int(gr) - 1) if gr.isdigit() else '?'
    return f"{gn} {gl}"

def grado_clave(sid):
    sid = str(sid).strip()
    return '10' if sid.startswith('10') else (sid[0] if sid else '?')

def nivel_desempeno(nota):
    if nota >= 90: return "Superior", "#1a7a1a"
    if nota >= 70: return "Alto",     "#2d7fcc"
    if nota >= 60: return "Basico",   "#d48a00"
    return "Bajo", "#cc2200"

def barra(valor, color="#4472C4", h=16):
    p = min(max(float(valor), 0), 100)
    return (f'<div style="background:#e8e8e8;border-radius:3px;height:{h}px;width:100%;">'
            f'<div style="background:{color};width:{p:.1f}%;height:{h}px;border-radius:3px;'
            f'display:flex;align-items:center;padding-left:5px;">'
            f'<span style="color:#fff;font-size:10px;font-weight:700;white-space:nowrap;">{p:.1f}</span>'
            f'</div></div>')

def badge(nivel):
    colores = {"Superior":"#1a7a1a","Alto":"#2d7fcc","Basico":"#d48a00","Bajo":"#cc2200"}
    c = colores.get(nivel, "#666")
    return f'<span style="background:{c};color:#fff;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:700;">{nivel}</span>'

def recomendacion(area, nota):
    nv, _ = nivel_desempeno(nota)
    tabla = {
        "Matematicas": {
            "Bajo":    "Necesita refuerzo urgente en operaciones numericas y resolucion de problemas. Se recomienda practica diaria y uso de material concreto.",
            "Basico":  "Comprende conceptos basicos. Debe fortalecer el razonamiento logico y la aplicacion en situaciones nuevas.",
            "Alto":    "Buen dominio matematico. Puede profundizar con problemas de mayor complejidad y proyectos estadisticos.",
            "Superior":"Desempeno sobresaliente. Se recomienda participar en olimpiadas matematicas.",
        },
        "Etica": {
            "Bajo":    "Necesita fortalecer la reflexion sobre valores y ciudadania. Se recomienda lectura de casos eticos y discusion grupal.",
            "Basico":  "Reconoce valores basicos. Debe trabajar la argumentacion y la toma de decisiones fundamentadas.",
            "Alto":    "Buen desempeno etico. Puede liderar dinamicas de convivencia y proyectos ciudadanos.",
            "Superior":"Excelente formacion ciudadana. Se recomienda como monitor de convivencia.",
        },
        "Lengua Castellana": {
            "Bajo":    "Dificultades en comprension lectora. Lectura diaria de textos cortos e identificacion de ideas principales.",
            "Basico":  "Comprende textos sencillos. Debe mejorar la lectura critica, inferencia y produccion escrita.",
            "Alto":    "Buena comprension textual. Puede leer textos literarios y academicos de mayor complejidad.",
            "Superior":"Excelente competencia comunicativa. Se recomienda participar en concursos de escritura y debate.",
        },
        "Ed. Fisica": {
            "Bajo":    "Reforzar conceptos de condicion fisica y habitos saludables. Participacion activa en todas las clases.",
            "Basico":  "Reconoce conceptos basicos. Debe mejorar habitos saludables y expresion corporal.",
            "Alto":    "Buen desempeno. Profundizar en rendimiento fisico y deporte formativo.",
            "Superior":"Sobresaliente. Se recomienda participar en equipos deportivos representativos.",
        },
        "Ciencias Sociales": {
            "Bajo":    "Reforzar procesos historicos y geograficos con mapas, lineas de tiempo y textos historicos.",
            "Basico":  "Comprende eventos basicos. Mejorar analisis de causas y consecuencias sociales.",
            "Alto":    "Buen manejo del area. Profundizar con investigacion sobre historia local y regional.",
            "Superior":"Excelente analisis social. Se recomienda participar en debates e investigacion historica.",
        },
        "Gestion": {
            "Bajo":    "Fortalecer competencias ciudadanas y pensamiento emprendedor con casos practicos.",
            "Basico":  "Reconoce conceptos basicos. Mejorar aplicacion de habilidades ciudadanas en situaciones reales.",
            "Alto":    "Buen desempeno. Puede proponer proyectos productivos escolares.",
            "Superior":"Excelente competencia emprendedora. Liderar proyectos de emprendimiento escolar.",
        },
        "Ciencias Naturales": {
            "Bajo":    "Reforzar conceptos biologicos con materiales visuales, experimentos sencillos y lectura cientifica.",
            "Basico":  "Comprende nociones basicas. Mejorar explicacion de fenomenos y metodo cientifico.",
            "Alto":    "Buen dominio. Profundizar con proyectos de investigacion cientifica escolar.",
            "Superior":"Desempeno cientifico sobresaliente. Participar en ferias de ciencias.",
        },
        "Artistica": {
            "Bajo":    "Reforzar apreciacion artistica con talleres de expresion y sensibilidad creativa.",
            "Basico":  "Reconoce conceptos basicos. Desarrollar produccion creativa y sensibilidad perceptiva.",
            "Alto":    "Buen desempeno. Profundizar en tecnicas de expresion y comprension critica del arte.",
            "Superior":"Excelente sensibilidad artistica. Participar en eventos culturales.",
        },
        "Ingles": {
            "Bajo":    "Practica diaria de vocabulario y ejercicios de lectura basica en ingles.",
            "Basico":  "Comprende estructuras basicas. Mejorar lectura y escritura de oraciones sencillas.",
            "Alto":    "Buen desempeno. Practicar con materiales autenticos: canciones, videos y textos.",
            "Superior":"Excelente competencia. Participar en inmersion y concursos de ingles.",
        },
        "Religion": {
            "Bajo":    "Fortalecer comprension de contenidos religiosos y antropologicos con reflexion y lectura.",
            "Basico":  "Reconoce contenidos basicos. Mejorar argumentacion y valoracion de temas eticos.",
            "Alto":    "Buen desempeno. Profundizar en dialogo interreligioso y valores universales.",
            "Superior":"Excelente comprension religiosa. Participar en proyectos de valores y pastoral.",
        },
        "Tecnologia": {
            "Bajo":    "Practica con herramientas digitales basicas y conceptos fundamentales de tecnologia.",
            "Basico":  "Reconoce conceptos tecnologicos. Mejorar uso y apropiacion de herramientas digitales.",
            "Alto":    "Buen manejo tecnologico. Profundizar con proyectos de diseno y resolucion de problemas.",
            "Superior":"Excelente desempeno. Participar en proyectos de innovacion y robotica escolar.",
        },
        "Filosofia": {
            "Bajo":    "Reforzar conceptos epistemologicos y gnoseologicos con lectura de textos filosoficos basicos.",
            "Basico":  "Comprende nociones basicas. Mejorar argumentacion y analisis critico de ideas filosoficas.",
            "Alto":    "Buen desempeno. Profundizar en corrientes filosoficas y su aplicacion al pensamiento critico.",
            "Superior":"Excelente formacion filosofica. Participar en debates academicos y olimpiadas de filosofia.",
        },
        "C. Politicas": {
            "Bajo":    "Reforzar conceptos de constitucion, derechos y mecanismos de participacion ciudadana.",
            "Basico":  "Reconoce conceptos basicos. Mejorar comprension de estructura del Estado y participacion politica.",
            "Alto":    "Buen desempeno. Profundizar en analisis de politicas publicas y ejercicio ciudadano.",
            "Superior":"Excelente comprension politica. Liderar proyectos de participacion y democracia escolar.",
        },
        "Quimica": {
            "Bajo":    "Reforzar conceptos basicos de estructura atomica, enlaces y reacciones quimicas con experimentos.",
            "Basico":  "Comprende nociones basicas. Mejorar formulacion quimica y comprension de reacciones.",
            "Alto":    "Buen dominio. Profundizar con problemas de estequiometria y quimica organica.",
            "Superior":"Desempeno sobresaliente. Participar en ferias de ciencias y olimpiadas de quimica.",
        },
        "Fisica": {
            "Bajo":    "Reforzar conceptos de cinematica, dinamica y energia con ejercicios practicos y videos.",
            "Basico":  "Comprende nociones basicas. Mejorar resolucion de problemas con formulas fisicas.",
            "Alto":    "Buen desempeno. Profundizar en fisica moderna y aplicaciones tecnologicas.",
            "Superior":"Excelente desempeno. Participar en olimpiadas de fisica y proyectos de ingenieria.",
        },
    }
    return tabla.get(area, {}).get(nv, "Continua con esfuerzo y dedicacion para mejorar tu desempeno.")

# ─── Leer tabla de respuestas del Excel ───────────────────────────────────────

def leer_tabla(excel_path):
    """Retorna dict: {num_pregunta: {area, topic, competencia}}"""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    area_actual = None
    tabla = {}
    for row in ws.iter_rows(values_only=True):
        if row[0] is not None and isinstance(row[0], str) and row[0] not in ('AREA',):
            area_actual = row[0].strip()
        q = row[1]
        if isinstance(q, int) and area_actual:
            topic = str(row[2]).strip() if row[2] else ""
            comp  = str(row[3]).strip() if row[3] else ""
            tabla[q] = {"area": area_actual, "topic": topic, "competencia": comp}
    return tabla

# ─── Cargar datos ─────────────────────────────────────────────────────────────

def _qnum(col_name):
    """Extrae el número de '#N Points Earned'."""
    m = re.search(r'#(\d+)', col_name)
    return int(m.group(1)) if m else None

def cargar_grado(f_s1, f_s2, areas_preguntas=None):
    if areas_preguntas is None:
        areas_preguntas = AREAS_PREGUNTAS

    s1 = f_s1 if isinstance(f_s1, pd.DataFrame) else pd.read_csv(BASE_DIR / f_s1)
    s2 = f_s2 if isinstance(f_s2, pd.DataFrame) else pd.read_csv(BASE_DIR / f_s2)

    cols_s1 = [c for c in s1.columns if c.startswith('#')]
    cols_s2 = [c for c in s2.columns if c.startswith('#')]
    offset  = len(cols_s1)

    # Usar el número real del encabezado para el mapeo correcto aunque haya saltos
    s1 = s1.rename(columns={c: f"Q{_qnum(c)}"          for c in cols_s1})
    s2 = s2.rename(columns={c: f"Q{offset + _qnum(c)}" for c in cols_s2})

    keys  = ['Student First Name', 'Student Last Name', 'Student ID']
    q1    = [f"Q{_qnum(c)}"          for c in cols_s1]
    q2    = [f"Q{offset + _qnum(c)}" for c in cols_s2]

    merged = pd.merge(s1[keys+q1], s2[keys+q2], on=keys, how='outer')
    merged['Student ID']  = merged['Student ID'].astype(str).str.strip()
    merged['Nombre']      = (merged['Student First Name'].str.strip() + ' '
                             + merged['Student Last Name'].str.strip())
    merged['Apellido']    = merged['Student Last Name'].str.strip()
    merged['NombrePila']  = merged['Student First Name'].str.strip()
    merged['Grupo']       = merged['Student ID'].apply(grupo_desde_id)
    merged['GradoClave']  = merged['Student ID'].apply(grado_clave)

    for area, qs in areas_preguntas.items():
        cols = [f"Q{q}" for q in qs if f"Q{q}" in merged.columns]
        if cols:
            merged[area] = (merged[cols].fillna(0).sum(axis=1) / len(cols)) * 100
        else:
            merged[area] = 0.0

    merged['Promedio'] = merged[list(areas_preguntas.keys())].mean(axis=1)
    merged['Pos_Grado'] = merged['Promedio'].rank(ascending=False, method='min').astype(int)
    merged['Pos_Grupo'] = merged.groupby('Grupo')['Promedio'].rank(ascending=False, method='min').astype(int)
    merged['Tam_Grupo'] = merged.groupby('Grupo')['Promedio'].transform('count').astype(int)
    return merged

# ─── Analisis por tema ────────────────────────────────────────────────────────

def analisis_temas(df_grupo, tabla, area, qs_area, colores_area=None):
    """Retorna HTML con analisis de temas para un grupo/grado."""
    topics = {}
    for q in qs_area:
        if q not in tabla: continue
        info = tabla[q]
        t = info['topic']
        col = f"Q{q}"
        if col not in df_grupo.columns: continue
        pct = df_grupo[col].fillna(0).mean() * 100
        if t not in topics:
            topics[t] = {"total": 0, "pct_sum": 0.0, "preguntas": []}
        topics[t]["total"] += 1
        topics[t]["pct_sum"] += pct
        topics[t]["preguntas"].append((q, pct))

    if not topics:
        return ""

    rows = ""
    for t, d in sorted(topics.items(), key=lambda x: -x[1]["pct_sum"]/x[1]["total"]):
        prom = d["pct_sum"] / d["total"]
        nv, _ = nivel_desempeno(prom)
        rows += f"""
        <tr>
          <td>{t}</td>
          <td>{d['total']}</td>
          <td>{barra(prom, (colores_area or COLORES_AREA).get(area,'#555'))}</td>
          <td style="text-align:center;font-weight:700;">{prom:.1f}</td>
          <td style="text-align:center;">{badge(nv)}</td>
        </tr>"""

    sorted_topics = sorted(topics.items(), key=lambda x: x[1]["pct_sum"]/x[1]["total"])
    peor  = sorted_topics[0][0]  if sorted_topics else "-"
    mejor = sorted_topics[-1][0] if sorted_topics else "-"

    return f"""
    <div style="margin:10px 0 4px 0;padding:8px 12px;background:#e3f0fb;border-left:4px solid #4472C4;border-radius:3px;font-size:12px;">
      &#x2714; <strong>Tema con mejor desempeno:</strong> {mejor} &nbsp;|&nbsp;
      &#x26A0; <strong>Tema a reforzar:</strong> {peor}
    </div>
    <table>
      <tr>
        <th>Tema</th><th style="width:6%;text-align:center;">Pregs.</th>
        <th>Desempeno</th><th style="width:10%;text-align:center;">Prom.</th>
        <th style="width:11%;text-align:center;">Nivel</th>
      </tr>
      {rows}
    </table>"""

# ─── CSS global ───────────────────────────────────────────────────────────────

def css():
    return """
<style>
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:Arial,'Helvetica Neue',sans-serif;font-size:13px;color:#222;background:#f0f2f5;}
.page{background:#fff;max-width:920px;margin:24px auto;padding:36px 44px;border-radius:8px;
      box-shadow:0 2px 12px rgba(0,0,0,.12);page-break-after:always;}
.header{display:flex;align-items:center;border-bottom:3px solid #1a3a6b;
        padding-bottom:12px;margin-bottom:16px;}
.header-ico{font-size:36px;margin-right:14px;}
.h1{font-size:15px;color:#1a3a6b;font-weight:700;}
.h2{font-size:12px;color:#555;margin-top:2px;}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:8px 24px;margin:12px 0 16px 0;}
.info-label{font-size:10px;color:#888;text-transform:uppercase;letter-spacing:.4px;}
.info-val{font-size:14px;font-weight:700;color:#1a3a6b;}
.sec{font-size:13px;font-weight:700;color:#1a3a6b;border-left:4px solid #1a3a6b;
     padding-left:8px;margin:18px 0 8px 0;}
.sumbox{background:#1a3a6b;color:#fff;border-radius:8px;padding:12px 16px;margin:12px 0;
        display:flex;justify-content:space-between;align-items:center;gap:8px;}
.snum{font-size:30px;font-weight:700;}
.slabel{font-size:11px;opacity:.75;}
table{width:100%;border-collapse:collapse;margin:8px 0;}
th{background:#1a3a6b;color:#fff;padding:7px 10px;text-align:left;font-size:11px;}
td{padding:6px 10px;border-bottom:1px solid #e8e8e8;font-size:12px;vertical-align:middle;}
tr:nth-child(even) td{background:#f7f8fc;}
.rec{background:#f4f7fb;border-left:4px solid #4472C4;padding:7px 11px;border-radius:3px;
     font-size:12px;margin:4px 0;}
.rec-area{font-weight:700;color:#1a3a6b;margin-bottom:2px;}
.alert-ok{background:#e8f5e9;border-left:4px solid #4caf50;padding:8px 12px;
          border-radius:3px;font-size:12px;color:#1b5e20;margin:6px 0;}
.alert-w{background:#fff8e1;border-left:4px solid #ffc107;padding:8px 12px;
         border-radius:3px;font-size:12px;color:#6d4200;margin:6px 0;}
.alert-i{background:#e3f0fb;border-left:4px solid #2196f3;padding:8px 12px;
         border-radius:3px;font-size:12px;color:#0d47a1;margin:6px 0;}
.foot{margin-top:20px;border-top:1px solid #ddd;padding-top:8px;
      font-size:10px;color:#999;text-align:center;}
@media print{body{background:#fff;}.page{box-shadow:none;margin:0;border-radius:0;}}
</style>"""

def encabezado(subtitulo=""):
    sub = f" &mdash; {subtitulo}" if subtitulo else ""
    return f"""
    <div class="header">
      <div class="header-ico">&#127979;</div>
      <div>
        <div class="h1">{COLEGIO}</div>
        <div class="h2">{EXAMEN} &mdash; {ANIO}{sub}</div>
      </div>
    </div>"""

def pie():
    return f'<div class="foot">{COLEGIO} &bull; {EXAMEN} {ANIO} &bull; Generado automaticamente</div>'

# ─── 1. Informes individuales (1 archivo por estudiante) ──────────────────────

def informes_individuales(df, grado_str, tabla, areas_preguntas=None):
    if areas_preguntas is None:
        areas_preguntas = AREAS_PREGUNTAS
    out = PORTAL_DIR / "informes" / grado_str / "individuales"
    out.mkdir(parents=True, exist_ok=True)
    total_grado = len(df)
    generados = []

    for _, row in df.iterrows():
        nombre   = row['Nombre']
        grupo    = row['Grupo']
        sid      = row['Student ID']
        prom     = row['Promedio']
        nv, _    = nivel_desempeno(prom)
        pg       = int(row['Pos_Grupo'])
        tg_g     = int(row['Tam_Grupo'])
        pg_grado = int(row['Pos_Grado'])

        # Tabla de notas + recomendaciones
        rows_notas = ""
        recs = ""
        for area in areas_preguntas:
            nota  = row[area]
            nva, nc = nivel_desempeno(nota)
            rows_notas += f"""
            <tr>
              <td><strong>{area}</strong></td>
              <td>{barra(nota, COLORES_AREA.get(area,'#555'))}</td>
              <td style="text-align:center;">{badge(nva)}</td>
              <td style="text-align:center;font-weight:700;">{nota:.1f}</td>
            </tr>"""
            if nota < 70:
                recs += f"""
                <div class="rec">
                  <div class="rec-area">&#128218; {area} ({nota:.1f}/100)</div>
                  {recomendacion(area, nota)}
                </div>"""

        if not recs:
            recs = '<div class="alert-ok">&#10003; Excelente desempeno en todas las areas. Continua con este nivel.</div>'

        html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Informe - {nombre}</title>
{css()}
</head>
<body>
<div class="page">
  {encabezado(f'Informe Individual &mdash; {grupo}')}
  <div class="grid2">
    <div><div class="info-label">Estudiante</div><div class="info-val">{nombre}</div></div>
    <div><div class="info-label">Grupo</div><div class="info-val">{grupo}</div></div>
    <div><div class="info-label">Codigo</div><div class="info-val">{sid}</div></div>
    <div><div class="info-label">Periodo</div><div class="info-val">Primer Periodo {ANIO}</div></div>
  </div>

  <div class="sumbox">
    <div>
      <div class="slabel">PROMEDIO GENERAL</div>
      <div class="snum">{prom:.1f}<span style="font-size:16px;">/100</span></div>
      {badge(nv)}
    </div>
    <div style="text-align:center;">
      <div class="slabel">Posicion en grupo</div>
      <div class="snum">{pg}<span style="font-size:14px;">/{tg_g}</span></div>
      <div class="slabel">{grupo}</div>
    </div>
    <div style="text-align:center;">
      <div class="slabel">Posicion en grado</div>
      <div class="snum">{pg_grado}<span style="font-size:14px;">/{total_grado}</span></div>
      <div class="slabel">{grupo.split()[0]}</div>
    </div>
  </div>

  <div class="sec">Resultados por Area</div>
  <table>
    <tr><th style="width:22%;">Area</th><th>Desempeno</th>
        <th style="width:12%;text-align:center;">Nivel</th>
        <th style="width:10%;text-align:center;">Nota</th></tr>
    {rows_notas}
  </table>

  <div class="sec">Recomendaciones Personalizadas</div>
  {recs}
  {pie()}
</div>
</body>
</html>"""

        fname = f"{sid}_{sanitize(row['Apellido'])}_{sanitize(row['NombrePila'])}.html"
        fpath = out / fname
        fpath.write_text(html, encoding='utf-8')
        generados.append({"id": sid, "nombre": nombre, "grupo": grupo,
                          "grado": grado_str, "archivo": f"informes/{grado_str}/individuales/{fname}"})

    print(f"  ✓ {len(generados)} informes individuales — {grado_str}")
    return generados

# ─── 2. Informes docentes (por area) ─────────────────────────────────────────

def informes_docentes(df, grado_str, tabla, areas_preguntas=None, colores_area=None):
    if areas_preguntas is None:
        areas_preguntas = AREAS_PREGUNTAS
    if colores_area is None:
        colores_area = COLORES_AREA
    out = PORTAL_DIR / "informes" / grado_str / "docentes"
    out.mkdir(parents=True, exist_ok=True)
    archivos = []
    grupos = sorted(df['Grupo'].unique())

    for area, qs_area in areas_preguntas.items():
        pages = []

        # ── Pagina resumen del grado ──
        prom_g = df[area].mean()
        total  = len(df)
        bajo    = (df[area] < 60).sum()
        basico  = ((df[area] >= 60) & (df[area] < 70)).sum()
        alto    = ((df[area] >= 70) & (df[area] < 90)).sum()
        superior= (df[area] >= 90).sum()

        rows_grupos = ""
        for g in grupos:
            sub = df[df['Grupo'] == g]
            pm = sub[area].mean(); n = len(sub)
            apro = (sub[area] >= 60).sum()
            nv, _ = nivel_desempeno(pm)
            rows_grupos += f"""
            <tr>
              <td><strong>{g}</strong></td>
              <td style="text-align:center;">{n}</td>
              <td>{barra(pm, colores_area.get(area,'#555'))}</td>
              <td style="text-align:center;font-weight:700;">{pm:.1f}</td>
              <td style="text-align:center;">{badge(nv)}</td>
              <td style="text-align:center;">{apro} ({apro/n*100:.0f}%)</td>
            </tr>"""

        top5 = df.nlargest(5, area)[['Nombre','Grupo',area]]
        bot5 = df.nsmallest(5, area)[['Nombre','Grupo',area]]
        r_top = "".join(f"<tr><td>{r['Nombre']}</td><td>{r['Grupo']}</td>"
                        f"<td style='text-align:center;font-weight:700;color:#1a7a1a;'>{r[area]:.1f}</td></tr>"
                        for _, r in top5.iterrows())
        r_bot = "".join(f"<tr><td>{r['Nombre']}</td><td>{r['Grupo']}</td>"
                        f"<td style='text-align:center;font-weight:700;color:#cc2200;'>{r[area]:.1f}</td></tr>"
                        for _, r in bot5.iterrows())

        # Analisis de temas
        html_temas = analisis_temas(df, tabla, area, qs_area, colores_area=colores_area)
        sec_temas = (f'<div class="sec">Analisis por Tema &mdash; {area} (Grado completo)</div>{html_temas}'
                     if html_temas else "")

        pages.append(f"""
        <div class="page">
          {encabezado(f'Informe Docente &mdash; {area} &mdash; {grado_str}')}
          <div class="sumbox">
            <div><div class="slabel">PROMEDIO GRADO</div>
                 <div class="snum">{prom_g:.1f}<span style="font-size:16px;">/100</span></div></div>
            <div style="text-align:center;"><div class="slabel">Evaluados</div>
                 <div class="snum">{total}</div></div>
            <div style="text-align:center;"><div class="slabel">Aprobados (&#8805;60)</div>
                 <div class="snum">{(df[area]>=60).sum()} <span style="font-size:14px;">({(df[area]>=60).sum()/total*100:.0f}%)</span></div></div>
          </div>
          <div class="sec">Resultados por Grupo</div>
          <table>
            <tr><th>Grupo</th><th style="text-align:center;">Est.</th>
                <th>Desempeno</th><th style="text-align:center;">Prom.</th>
                <th style="text-align:center;">Nivel</th><th style="text-align:center;">Aprobados</th></tr>
            {rows_grupos}
          </table>
          <div class="sec">Distribucion por Nivel</div>
          <table>
            <tr><th>Nivel</th><th style="text-align:center;">Cantidad</th>
                <th style="text-align:center;">Porcentaje</th><th>Rango</th></tr>
            <tr><td>{badge('Superior')}</td><td style="text-align:center;">{superior}</td>
                <td style="text-align:center;">{superior/total*100:.1f}%</td><td>90 &ndash; 100</td></tr>
            <tr><td>{badge('Alto')}</td><td style="text-align:center;">{alto}</td>
                <td style="text-align:center;">{alto/total*100:.1f}%</td><td>70 &ndash; 89.9</td></tr>
            <tr><td>{badge('Basico')}</td><td style="text-align:center;">{basico}</td>
                <td style="text-align:center;">{basico/total*100:.1f}%</td><td>60 &ndash; 69.9</td></tr>
            <tr><td>{badge('Bajo')}</td><td style="text-align:center;">{bajo}</td>
                <td style="text-align:center;">{bajo/total*100:.1f}%</td><td>0 &ndash; 59.9</td></tr>
          </table>
          {sec_temas}
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:12px;">
            <div><div class="sec" style="margin-top:0;">&#127942; Mejores puntajes</div>
              <table><tr><th>Estudiante</th><th>Grupo</th>
                  <th style="text-align:center;">Nota</th></tr>{r_top}</table></div>
            <div><div class="sec" style="margin-top:0;">&#128204; Atencion prioritaria</div>
              <table><tr><th>Estudiante</th><th>Grupo</th>
                  <th style="text-align:center;">Nota</th></tr>{r_bot}</table></div>
          </div>
          {pie()}
        </div>""")

        # ── Pagina por grupo ──
        for g in grupos:
            sub = df[df['Grupo'] == g].sort_values(area, ascending=False).reset_index(drop=True)
            html_temas_g = analisis_temas(sub, tabla, area, qs_area, colores_area=colores_area)
            sec_temas_g  = (f'<div class="sec">Analisis por Tema &mdash; {g}</div>{html_temas_g}'
                            if html_temas_g else "")

            rows_est = ""
            for i, (_, r) in enumerate(sub.iterrows()):
                nva, _ = nivel_desempeno(r[area])
                rows_est += f"""
                <tr>
                  <td style="text-align:center;">{i+1}</td>
                  <td><strong>{r['Nombre']}</strong></td>
                  <td>{barra(r[area], colores_area.get(area,'#555'))}</td>
                  <td style="text-align:center;font-weight:700;">{r[area]:.1f}</td>
                  <td style="text-align:center;">{badge(nva)}</td>
                </tr>"""

            pages.append(f"""
            <div class="page">
              {encabezado(f'Informe Docente &mdash; {area} &mdash; {g}')}
              <div class="grid2">
                <div><div class="info-label">Promedio del grupo</div>
                     <div class="info-val">{sub[area].mean():.1f}</div></div>
                <div><div class="info-label">Estudiantes</div>
                     <div class="info-val">{len(sub)}</div></div>
                <div><div class="info-label">Nota maxima</div>
                     <div class="info-val">{sub[area].max():.1f}</div></div>
                <div><div class="info-label">Nota minima</div>
                     <div class="info-val">{sub[area].min():.1f}</div></div>
              </div>
              {sec_temas_g}
              <div class="sec">Ranking &mdash; {g}</div>
              <table>
                <tr><th style="width:5%;text-align:center;">#</th><th>Estudiante</th>
                    <th>Desempeno</th><th style="width:10%;text-align:center;">Nota</th>
                    <th style="width:12%;text-align:center;">Nivel</th></tr>
                {rows_est}
              </table>
              {pie()}
            </div>""")

        fname = f"{sanitize(area)}_{grado_str}.html"
        fpath = out / fname
        html  = f'<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Docente {area} {grado_str}</title>{css()}</head><body>'
        html += "\n".join(pages) + "</body></html>"
        fpath.write_text(html, encoding='utf-8')
        archivos.append({"area": area, "grado": grado_str,
                         "archivo": f"informes/{grado_str}/docentes/{fname}"})

    print(f"  ✓ {len(archivos)} informes docentes — {grado_str}")
    return archivos

# ─── 3. Informes directores de grupo ─────────────────────────────────────────

def informes_directores(df, grado_str, areas_preguntas=None, colores_area=None):
    if areas_preguntas is None:
        areas_preguntas = AREAS_PREGUNTAS
    if colores_area is None:
        colores_area = COLORES_AREA
    out = PORTAL_DIR / "informes" / grado_str / "directores"
    out.mkdir(parents=True, exist_ok=True)
    grupos   = sorted(df['Grupo'].unique())
    archivos = []

    for grupo in grupos:
        sub = df[df['Grupo'] == grupo].copy()
        sub['Pos_Loc'] = sub['Promedio'].rank(ascending=False, method='min').astype(int)
        sub = sub.sort_values('Pos_Loc')

        n         = len(sub)
        prom_g    = sub['Promedio'].mean()
        prom_grado= df['Promedio'].mean()
        diff      = prom_g - prom_grado
        arrow     = ('&#9650; por encima' if diff >= 0 else '&#9660; por debajo') + f' ({abs(diff):.1f} pts)'
        acolor    = '#1a7a1a' if diff >= 0 else '#cc2200'

        bajo    = (sub['Promedio'] < 60).sum()
        basico  = ((sub['Promedio'] >= 60) & (sub['Promedio'] < 70)).sum()
        alto    = ((sub['Promedio'] >= 70) & (sub['Promedio'] < 90)).sum()
        superior= (sub['Promedio'] >= 90).sum()

        rows_est = ""
        for _, r in sub.iterrows():
            nv, _ = nivel_desempeno(r['Promedio'])
            pos_grado_est = int(df['Promedio'].rank(ascending=False, method='min').loc[r.name])
            rows_est += f"""
            <tr>
              <td style="text-align:center;">{int(r['Pos_Loc'])}</td>
              <td><strong>{r['Nombre']}</strong></td>
              <td>{barra(r['Promedio'])}</td>
              <td style="text-align:center;font-weight:700;">{r['Promedio']:.1f}</td>
              <td style="text-align:center;">{badge(nv)}</td>
              <td style="text-align:center;">{pos_grado_est}/{len(df)}</td>
            </tr>"""

        rows_areas = ""
        for area in areas_preguntas:
            pm_a = sub[area].mean()
            pm_g = df[area].mean()
            d    = pm_a - pm_g
            ar   = f'<span style="color:{"#1a7a1a" if d>=0 else "#cc2200"}">{"&#9650;" if d>=0 else "&#9660;"} {abs(d):.1f}</span>'
            nv, _= nivel_desempeno(pm_a)
            rows_areas += f"""
            <tr>
              <td><strong>{area}</strong></td>
              <td>{barra(pm_a, colores_area.get(area,'#555'))}</td>
              <td style="text-align:center;font-weight:700;">{pm_a:.1f}</td>
              <td style="text-align:center;">{pm_g:.1f}</td>
              <td style="text-align:center;">{ar}</td>
              <td style="text-align:center;">{badge(nv)}</td>
            </tr>"""

        # Desglose: encabezados de areas para la tabla consolidada
        th_areas_desglose = "".join(
            f'<th style="text-align:center;background:{colores_area.get(area,"#555")};color:#fff;'
            f'font-size:10px;padding:5px 3px;min-width:55px;">{area}</th>'
            for area in areas_preguntas
        )
        bg_nivel = {"Superior":"#e8f5e9","Alto":"#e3f0fa","Basico":"#fff8e1","Bajo":"#ffebee"}
        fg_nivel = {"Superior":"#1a7a1a","Alto":"#2d7fcc","Basico":"#d48a00","Bajo":"#cc2200"}

        rows_desglose = ""
        for _, r in sub.iterrows():
            celdas_areas = ""
            for area in areas_preguntas:
                nota_a = r[area]
                nv_a, _ = nivel_desempeno(nota_a)
                bg = bg_nivel.get(nv_a, "#fff")
                fg = fg_nivel.get(nv_a, "#333")
                celdas_areas += (
                    f'<td style="text-align:center;background:{bg};color:{fg};'
                    f'font-weight:700;font-size:12px;">{nota_a:.1f}</td>'
                )
            rows_desglose += f"""
            <tr>
              <td style="text-align:center;">{int(r['Pos_Loc'])}</td>
              <td style="white-space:nowrap;"><strong>{r['Nombre']}</strong></td>
              <td style="text-align:center;font-weight:700;">{r['Promedio']:.1f}</td>
              {celdas_areas}
            </tr>"""

        sub['n_bajo'] = sub[list(areas_preguntas.keys())].apply(lambda r: (r < 60).sum(), axis=1)
        riesgo = sub[sub['n_bajo'] >= 3].sort_values('n_bajo', ascending=False)
        if len(riesgo):
            rows_r = "".join(
                f"<tr><td><strong>{r['Nombre']}</strong></td>"
                f"<td style='text-align:center;'>{int(r['n_bajo'])}</td>"
                f"<td style='font-size:11px;'>{', '.join(a for a in areas_preguntas if r[a]<60)}</td>"
                f"<td style='text-align:center;font-weight:700;color:#cc2200;'>{r['Promedio']:.1f}</td></tr>"
                for _, r in riesgo.iterrows()
            )
            sec_riesgo = f"""
            <div class="alert-w">&#9888; <strong>{len(riesgo)} estudiante(s)</strong> con Bajo en 3 o mas areas. Seguimiento prioritario recomendado.</div>
            <table><tr><th>Estudiante</th><th style="text-align:center;">Areas en Bajo</th>
                <th>Areas criticas</th><th style="text-align:center;">Promedio</th></tr>{rows_r}</table>"""
        else:
            sec_riesgo = '<div class="alert-ok">&#10003; Ningun estudiante con Bajo en 3 o mas areas.</div>'

        promedios = {a: sub[a].mean() for a in areas_preguntas}
        mejor_a = max(promedios, key=promedios.get)
        peor_a  = min(promedios, key=promedios.get)

        html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Director de Grupo &mdash; {grupo}</title>
{css()}
</head>
<body>
<div class="page">
  {encabezado(f'Informe Director de Grupo &mdash; {grupo}')}
  <div class="grid2">
    <div><div class="info-label">Grupo</div><div class="info-val">{grupo}</div></div>
    <div><div class="info-label">Total estudiantes</div><div class="info-val">{n}</div></div>
    <div><div class="info-label">Grado</div><div class="info-val">{grado_str}</div></div>
    <div><div class="info-label">Periodo</div><div class="info-val">Primer Periodo {ANIO}</div></div>
  </div>
  <div class="sumbox">
    <div><div class="slabel">PROMEDIO DEL GRUPO</div>
         <div class="snum">{prom_g:.1f}<span style="font-size:16px;">/100</span></div></div>
    <div style="text-align:center;"><div class="slabel">vs. promedio grado</div>
         <div class="snum">{prom_grado:.1f}</div>
         <div style="font-size:11px;color:{acolor};">{arrow}</div></div>
    <div style="text-align:center;"><div class="slabel">Aprobados (&#8805;60)</div>
         <div class="snum">{(sub['Promedio']>=60).sum()}<span style="font-size:14px;">/{n}</span></div></div>
  </div>
  <div class="sec">Distribucion por Nivel</div>
  <table>
    <tr><th>Nivel</th><th style="text-align:center;">Cant.</th>
        <th style="text-align:center;">%</th><th>Rango</th></tr>
    <tr><td>{badge('Superior')}</td><td style="text-align:center;">{superior}</td>
        <td style="text-align:center;">{superior/n*100:.1f}%</td><td>90 &ndash; 100</td></tr>
    <tr><td>{badge('Alto')}</td><td style="text-align:center;">{alto}</td>
        <td style="text-align:center;">{alto/n*100:.1f}%</td><td>70 &ndash; 89.9</td></tr>
    <tr><td>{badge('Basico')}</td><td style="text-align:center;">{basico}</td>
        <td style="text-align:center;">{basico/n*100:.1f}%</td><td>60 &ndash; 69.9</td></tr>
    <tr><td>{badge('Bajo')}</td><td style="text-align:center;">{bajo}</td>
        <td style="text-align:center;">{bajo/n*100:.1f}%</td><td>0 &ndash; 59.9</td></tr>
  </table>
  <div class="sec">Analisis por Area</div>
  <div class="alert-i">&#128170; Fortaleza: <strong>{mejor_a}</strong> ({promedios[mejor_a]:.1f}/100) &nbsp;&nbsp;
       &#128202; A reforzar: <strong>{peor_a}</strong> ({promedios[peor_a]:.1f}/100)</div>
  <table>
    <tr><th>Area</th><th>Desempeno grupo</th>
        <th style="text-align:center;">Prom. grupo</th>
        <th style="text-align:center;">Prom. grado</th>
        <th style="text-align:center;">Diferencia</th>
        <th style="text-align:center;">Nivel</th></tr>
    {rows_areas}
  </table>
  <div class="sec">Estudiantes en Seguimiento Prioritario</div>
  {sec_riesgo}
  {pie()}
</div>
<div class="page">
  {encabezado(f'Ranking Completo &mdash; {grupo}')}
  <div class="sec">Listado por Promedio General</div>
  <table>
    <tr><th style="text-align:center;width:5%;">#</th><th>Estudiante</th>
        <th>Promedio general</th><th style="text-align:center;width:10%;">Nota</th>
        <th style="text-align:center;width:12%;">Nivel</th>
        <th style="text-align:center;width:14%;">Pos. en grado</th></tr>
    {rows_est}
  </table>
  {pie()}
</div>
<div class="page" style="max-width:1200px;">
  {encabezado(f'Consolidado por Areas &mdash; {grupo}')}
  <div class="sec">Notas de todos los estudiantes por area &mdash; ordenadas de mayor a menor promedio</div>
  <div style="font-size:11px;margin-bottom:8px;display:flex;gap:12px;flex-wrap:wrap;">
    <span style="background:#e8f5e9;color:#1a7a1a;padding:2px 8px;border-radius:10px;font-weight:700;">Superior &#8805;90</span>
    <span style="background:#e3f0fa;color:#2d7fcc;padding:2px 8px;border-radius:10px;font-weight:700;">Alto 70&ndash;89</span>
    <span style="background:#fff8e1;color:#d48a00;padding:2px 8px;border-radius:10px;font-weight:700;">Basico 60&ndash;69</span>
    <span style="background:#ffebee;color:#cc2200;padding:2px 8px;border-radius:10px;font-weight:700;">Bajo &lt;60</span>
  </div>
  <div style="overflow-x:auto;">
  <table style="font-size:12px;width:100%;">
    <tr>
      <th style="text-align:center;width:3%;">#</th>
      <th style="min-width:160px;">Estudiante</th>
      <th style="text-align:center;min-width:60px;">Promedio</th>
      {th_areas_desglose}
    </tr>
    {rows_desglose}
  </table>
  </div>
  {pie()}
</div>
</body>
</html>"""

        fname = f"Director_{sanitize(grupo)}.html"
        (out / fname).write_text(html, encoding='utf-8')
        archivos.append({"grupo": grupo, "grado": grado_str,
                         "archivo": f"informes/{grado_str}/directores/{fname}"})

    print(f"  ✓ {len(archivos)} informes directores — {grado_str}")
    return archivos

# ─── 4. Desglose por pregunta ────────────────────────────────────────────────

def desglose_preguntas(df, grado_str, tabla, areas_preguntas=None, colores_area=None):
    """
    Genera un HTML por área con tabla pregunta×grupo:
    columnas = grupos, filas = preguntas (con tema), celdas = aciertos/total + barra.
    Salida: portal/informes/<grado>/desglose/<area>.html
    """
    if areas_preguntas is None:
        areas_preguntas = AREAS_PREGUNTAS
    if colores_area is None:
        colores_area = COLORES_AREA

    out = PORTAL_DIR / "informes" / grado_str / "desglose"
    out.mkdir(parents=True, exist_ok=True)

    grupos = sorted(df['Grupo'].unique())
    archivos = []

    for area, qs_area in areas_preguntas.items():
        color = colores_area.get(area, "#4472C4")
        qs_presentes = [q for q in qs_area if f"Q{q}" in df.columns]
        if not qs_presentes:
            continue

        # ── Encabezado de tabla: una columna por grupo + totales ──
        th_grupos = "".join(
            f'<th colspan="2" style="text-align:center;background:{color};color:#fff;">{g}</th>'
            for g in grupos
        )
        th_sub = "".join(
            '<th style="text-align:center;font-size:10px;color:#555;">Acier.</th>'
            '<th style="text-align:center;font-size:10px;color:#555;">%</th>'
            for _ in grupos
        )

        # ── Filas: una por pregunta ──
        filas = ""
        for q in qs_presentes:
            col = f"Q{q}"
            info = tabla.get(q, {})
            topic = info.get("topic", "—")
            comp  = info.get("competencia", "")

            celdas = ""
            for g in grupos:
                sub = df[df['Grupo'] == g]
                n   = len(sub)
                if n == 0:
                    celdas += '<td style="text-align:center;">—</td><td>—</td>'
                    continue
                ok  = int(sub[col].fillna(0).sum())
                pct = ok / n * 100
                # color de celda según porcentaje
                if pct >= 70:
                    bg = "#e8f5e9"; fg = "#1a7a1a"
                elif pct >= 50:
                    bg = "#fff8e1"; fg = "#d48a00"
                else:
                    bg = "#ffebee"; fg = "#cc2200"
                bar_w = f"{pct:.0f}%"
                celdas += (
                    f'<td style="text-align:center;background:{bg};color:{fg};font-weight:700;">{ok}/{n}</td>'
                    f'<td style="padding:4px 6px;background:{bg};">'
                    f'<div style="background:#ddd;border-radius:3px;height:12px;width:100%;">'
                    f'<div style="background:{color};width:{bar_w};height:12px;border-radius:3px;"></div>'
                    f'</div>'
                    f'<span style="font-size:10px;color:{fg};font-weight:700;">{pct:.0f}%</span>'
                    f'</td>'
                )

            # fila total grado
            n_tot = len(df)
            ok_tot = int(df[col].fillna(0).sum())
            pct_tot = ok_tot / n_tot * 100 if n_tot else 0
            nv_tot, _ = nivel_desempeno(pct_tot)
            tooltip = f'title="{comp}"' if comp else ""
            filas += (
                f'<tr>'
                f'<td style="text-align:center;font-weight:700;color:{color};">P{q}</td>'
                f'<td {tooltip} style="max-width:220px;white-space:nowrap;overflow:hidden;'
                f'text-overflow:ellipsis;font-size:12px;">{topic}</td>'
                f'{celdas}'
                f'<td style="text-align:center;font-weight:700;">{ok_tot}/{n_tot}</td>'
                f'<td style="text-align:center;">{badge(nv_tot)}</td>'
                f'</tr>'
            )

        # ── Tabla resumen por grupo (promedio del área) ──
        resumen = ""
        for g in grupos:
            sub = df[df['Grupo'] == g]
            if len(sub) == 0:
                continue
            prom = sub[area].mean()
            nv, _ = nivel_desempeno(prom)
            apro = (sub[area] >= 60).sum()
            resumen += (
                f'<tr>'
                f'<td><strong>{g}</strong></td>'
                f'<td style="text-align:center;">{len(sub)}</td>'
                f'<td>{barra(prom, color)}</td>'
                f'<td style="text-align:center;font-weight:700;">{prom:.1f}</td>'
                f'<td style="text-align:center;">{badge(nv)}</td>'
                f'<td style="text-align:center;">{apro} ({apro/len(sub)*100:.0f}%)</td>'
                f'</tr>'
            )

        html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Desglose por Pregunta — {area} — {grado_str}</title>
{css()}
<style>
body {{ font-family:'Segoe UI',sans-serif; background:#f4f6fb; padding:20px; }}
h1 {{ color:{color}; font-size:20px; margin-bottom:4px; }}
h2 {{ color:#555; font-size:13px; font-weight:400; margin-bottom:20px; }}
.bloque {{ background:#fff; border-radius:12px; padding:20px;
           box-shadow:0 2px 12px rgba(0,0,0,.08); margin-bottom:28px; overflow-x:auto; }}
.sec {{ font-size:13px; font-weight:700; color:{color};
        border-left:4px solid {color}; padding-left:9px; margin:16px 0 10px; }}
table {{ border-collapse:collapse; width:100%; font-size:12px; min-width:500px; }}
th, td {{ border:1px solid #e0e0e0; padding:6px 8px; vertical-align:middle; }}
th {{ background:#f0f4fb; color:#2c3e50; font-size:11px; }}
tr:nth-child(even) td {{ background:#fafbff; }}
.legend {{ display:flex; gap:16px; font-size:11px; margin-top:8px; }}
.leg {{ display:flex; align-items:center; gap:4px; }}
.dot {{ width:12px; height:12px; border-radius:50%; }}
</style>
</head>
<body>
<h1>&#128202; Desglose por Pregunta &mdash; {area}</h1>
<h2>{COLEGIO} &bull; {grado_str} &bull; {EXAMEN} {ANIO}</h2>

<div class="bloque">
  <div class="sec">Resumen por Grupo</div>
  <table>
    <tr>
      <th>Grupo</th><th style="text-align:center;">Est.</th>
      <th>Desempeno</th><th style="text-align:center;">Promedio</th>
      <th style="text-align:center;">Nivel</th><th style="text-align:center;">Aprobados</th>
    </tr>
    {resumen}
  </table>
</div>

<div class="bloque">
  <div class="sec">Aciertos y Fallas por Pregunta</div>
  <div class="legend">
    <div class="leg"><div class="dot" style="background:#1a7a1a;"></div> &ge;70% (Bueno)</div>
    <div class="leg"><div class="dot" style="background:#d48a00;"></div> 50-69% (Regular)</div>
    <div class="leg"><div class="dot" style="background:#cc2200;"></div> &lt;50% (Critico)</div>
  </div>
  <br>
  <table>
    <tr>
      <th style="width:45px;">Preg.</th>
      <th>Tema / Competencia</th>
      {th_grupos}
      <th colspan="2" style="text-align:center;background:#2c3e50;color:#fff;">Grado Total</th>
    </tr>
    <tr>
      <th></th><th></th>
      {th_sub}
      <th style="text-align:center;font-size:10px;color:#555;">Acier.</th>
      <th style="text-align:center;font-size:10px;color:#555;">Nivel</th>
    </tr>
    {filas}
  </table>
</div>

{pie()}
</body>
</html>"""

        fname = f"desglose_{sanitize(area)}_{grado_str}.html"
        fpath = out / fname
        fpath.write_text(html, encoding='utf-8')
        archivos.append({"area": area, "grado": grado_str,
                         "archivo": f"informes/{grado_str}/desglose/{fname}"})

    print(f"  ✓ {len(archivos)} desgloses por pregunta — {grado_str}")
    return archivos


# ─── 5. Portal web ────────────────────────────────────────────────────────────

GRADE_NUM = {"Sexto":"6","Septimo":"7","Octavo":"8","Noveno":"9","Decimo":"10"}

def generar_portal(datos_grados, archivos_docentes, archivos_directores, archivos_desglose=None, areas_por_grado=None):
    """datos_grados: {grado_str: DataFrame}"""
    if archivos_desglose is None:
        archivos_desglose = []
    if areas_por_grado is None:
        areas_por_grado = {g: list(AREAS_PREGUNTAS.keys()) for g in datos_grados}

    # Construir JSON de estudiantes
    estudiantes_json = []
    for grado_str, df in datos_grados.items():
        ap = areas_por_grado.get(grado_str, list(AREAS_PREGUNTAS.keys()))
        for _, row in df.iterrows():
            notas = {a: round(float(row[a]), 1) for a in ap}
            fname = f"{row['Student ID']}_{sanitize(row['Apellido'])}_{sanitize(row['NombrePila'])}.html"
            estudiantes_json.append({
                "id":       row['Student ID'],
                "nombre":   row['Nombre'],
                "grupo":    row['Grupo'],
                "grado":    grado_str,
                "promedio": round(float(row['Promedio']), 1),
                "notas":    notas,
                "archivo":  f"informes/{grado_str}/individuales/{fname}",
            })

    grupos_por_grado = {}
    for g, df in datos_grados.items():
        grupos_por_grado[g] = sorted(df['Grupo'].unique().tolist())

    # Escribir data.js separado
    data_js = (f"const ESTUDIANTES = {json.dumps(estudiantes_json, ensure_ascii=False)};\n\n"
               f"const GRUPOS_POR_GRADO = {json.dumps(grupos_por_grado, ensure_ascii=False)};\n")
    (PORTAL_DIR / "data.js").write_text(data_js, encoding='utf-8')

    grados_sorted = sorted(datos_grados.keys())
    areas_por_grado_js = {g: list(areas) for g, areas in areas_por_grado.items()}

    # Opciones de grado para busqueda
    opciones_buscar = "\n".join(
        f'<option value="{g}">{g} ({GRADE_NUM.get(g,"?")}°)</option>'
        for g in grados_sorted
    )

    # Opciones de grado para descarga
    opciones_dl = "\n".join(
        f'<option value="{g}">{g}</option>'
        for g in grados_sorted
    )

    # Bloques de grupos por directores
    rep_grupos = ""
    for g in grados_sorted:
        dirs_g = [a for a in archivos_directores if a["grado"] == g]
        if dirs_g:
            btns = "".join(
                f'<a href="{a["archivo"]}" target="_blank" class="rep-btn">'
                f'<span class="rep-ico">&#128101;</span>{a["grupo"]}</a>'
                for a in dirs_g
            )
            rep_grupos += (f'<div class="rep-group"><div class="rep-group-title">{g}</div>'
                           f'<div class="rep-grid">{btns}</div></div>')

    # Bloques de docentes por area
    rep_docentes = ""
    for g in grados_sorted:
        docs_g = [a for a in archivos_docentes if a["grado"] == g]
        if docs_g:
            btns = "".join(
                f'<a href="{a["archivo"]}" target="_blank" class="rep-btn">'
                f'<span class="rep-ico">&#128218;</span>{a["area"]}</a>'
                for a in docs_g
            )
            rep_docentes += (f'<div class="rep-group"><div class="rep-group-title">{g}</div>'
                             f'<div class="rep-grid">{btns}</div></div>')

    # Bloques de desglose por pregunta
    rep_desglose = ""
    for g in grados_sorted:
        desg_g = [a for a in archivos_desglose if a["grado"] == g]
        if desg_g:
            btns = "".join(
                f'<a href="{a["archivo"]}" target="_blank" class="rep-btn">'
                f'<span class="rep-ico">&#128202;</span>{a["area"]}</a>'
                for a in desg_g
            )
            rep_desglose += (f'<div class="rep-group"><div class="rep-group-title">{g}</div>'
                             f'<div class="rep-grid">{btns}</div></div>')

    portal = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Portal Educativo &mdash; {COLEGIO}</title>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{
  font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
  background: linear-gradient(135deg, #1a3a6b 0%, #2d6ea8 100%);
  min-height: 100vh;
  display: flex;
  flex-direction: column;
  align-items: center;
  padding: 24px 16px 0 16px;
}}
.container {{
  background: #fff;
  border-radius: 20px;
  box-shadow: 0 20px 60px rgba(0,0,0,.35);
  width: 100%;
  max-width: 600px;
  padding: 36px 36px 28px 36px;
  animation: slideIn .45s ease-out;
}}
@keyframes slideIn {{ from{{opacity:0;transform:translateY(-24px)}} to{{opacity:1;transform:translateY(0)}} }}
.logo {{
  width: 72px; height: 72px;
  background: linear-gradient(135deg,#1a3a6b,#2d6ea8);
  border-radius: 50%;
  margin: 0 auto 16px;
  display: flex; align-items: center; justify-content: center;
  font-size: 34px; color: #fff;
  box-shadow: 0 8px 24px rgba(26,58,107,.3);
}}
.header {{ text-align:center; margin-bottom:26px; }}
.header h1 {{ color:#1a3a6b; font-size:21px; font-weight:700; line-height:1.3; }}
.header h2 {{ color:#7f8c8d; font-size:13px; margin-top:6px; }}
.tabs {{
  display:flex; gap:8px;
  background:#f0f2f8; border-radius:14px; padding:5px; margin-bottom:24px;
}}
.tab-btn {{
  flex:1; padding:11px 6px; border:none; border-radius:10px;
  font-size:13px; font-weight:700; cursor:pointer; transition:.2s;
  background:transparent; color:#7f8c8d; line-height:1.3;
}}
.tab-btn.active {{
  background:#fff; color:#1a3a6b;
  box-shadow: 0 4px 14px rgba(0,0,0,.12);
}}
.section {{ display:none; }}
.section.active {{ display:block; }}
.info-box {{
  background:#e8f0fb; border-left:4px solid #1a3a6b;
  border-radius:8px; padding:11px 14px;
  font-size:13px; color:#1a3a6b; margin-bottom:20px; line-height:1.5;
}}
label.field-label {{ display:block; font-weight:600; color:#2c3e50; font-size:13px; margin-bottom:6px; }}
select, input[type=text] {{
  width:100%; padding:13px 14px; border:2px solid #ecf0f1;
  border-radius:10px; font-size:15px; outline:none; transition:.25s;
  font-family: inherit; background:#fff; color:#2c3e50;
}}
select:focus, input[type=text]:focus {{ border-color:#1a3a6b; box-shadow:0 0 0 3px rgba(26,58,107,.1); }}
.field {{ margin-bottom:16px; }}
.search-btn {{
  width:100%; padding:15px; color:#fff; border:none; border-radius:10px;
  font-size:16px; font-weight:700; cursor:pointer; transition:.25s;
  background: linear-gradient(135deg,#1a3a6b,#2d6ea8);
  box-shadow: 0 8px 24px rgba(26,58,107,.3);
}}
.search-btn:hover {{ transform:translateY(-2px); box-shadow:0 12px 28px rgba(26,58,107,.35); }}
.search-btn:active {{ transform:translateY(0); }}
#resultados {{ margin-top:20px; }}
.result-card {{
  background:#f4f7fb;
  border-left:5px solid #1a3a6b;
  border-radius:10px; padding:14px 16px; margin-bottom:10px;
  display:flex; justify-content:space-between; align-items:center; gap:12px;
  cursor:pointer; transition:.2s; text-decoration:none;
}}
.result-card:hover {{ transform:translateX(4px); box-shadow:0 4px 18px rgba(0,0,0,.1); }}
.r-nombre {{ font-size:15px; font-weight:700; color:#1a3a6b; }}
.r-info {{ font-size:12px; color:#7f8c8d; margin-top:3px; }}
.r-nota {{ font-size:24px; font-weight:700; color:#1a3a6b; min-width:56px; text-align:center; }}
.r-nota small {{ display:block; font-size:11px; font-weight:600; }}
.badge {{ display:inline-block; padding:2px 8px; border-radius:10px; font-size:10px; font-weight:700; color:#fff; }}
.b-s{{background:#1a7a1a}} .b-a{{background:#2d7fcc}} .b-b{{background:#d48a00}} .b-l{{background:#c00}}
.no-result {{ text-align:center; padding:28px 0; color:#aaa; font-size:14px; }}
.error-box {{
  background:#fdecea; border-left:4px solid #e53935;
  border-radius:8px; padding:11px 14px; font-size:13px; color:#b71c1c;
  margin-top:14px; display:none;
}}
.rep-group {{ margin-bottom:20px; }}
.rep-group-title {{ font-size:13px; font-weight:700; color:#888; text-transform:uppercase;
                    letter-spacing:.5px; margin-bottom:10px; }}
.rep-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:8px; }}
.rep-btn {{
  display:flex; align-items:center; gap:8px;
  padding:11px 13px; background:#f4f7fb;
  border:1.5px solid #dde4f0; border-radius:10px;
  color:#1a3a6b; font-size:13px; font-weight:600;
  text-decoration:none; transition:.18s;
}}
.rep-btn:hover {{ background:#1a3a6b; color:#fff; border-color:#1a3a6b; }}
.rep-ico {{ font-size:16px; }}
.rep-section-title {{
  font-size:14px; font-weight:700; color:#1a3a6b;
  border-left:4px solid #1a3a6b; padding-left:9px; margin:0 0 14px 0;
}}
.subtabs {{
  display:flex; gap:6px;
  background:#f0f2f8; border-radius:12px; padding:4px; margin-bottom:20px;
}}
.subtab-btn {{
  flex:1; padding:9px 6px; border:none; border-radius:9px;
  font-size:13px; font-weight:700; cursor:pointer; transition:.2s;
  background:transparent; color:#7f8c8d;
}}
.subtab-btn.active {{
  background:#fff; color:#1a3a6b;
  box-shadow: 0 3px 10px rgba(0,0,0,.12);
}}
.subsection {{ display:none; }}
.subsection.active {{ display:block; }}
.dl-btn {{
  width:100%; padding:15px; color:#fff; border:none; border-radius:10px;
  font-size:16px; font-weight:700; cursor:pointer; transition:.25s;
  background: linear-gradient(135deg,#1a7a1a,#2e9c2e);
  box-shadow: 0 8px 24px rgba(26,122,26,.25); margin-top:6px;
}}
.dl-btn:hover {{ transform:translateY(-2px); }}
.dl-btn:active {{ transform:translateY(0); }}
footer {{
  width:100%; max-width:600px; text-align:center;
  padding:18px 0 22px 0; color:rgba(255,255,255,.75); font-size:12px; margin-top:14px;
}}
footer strong {{ display:block; font-size:14px; color:#fff; letter-spacing:1px; margin-top:4px; }}
/* ── Modal contraseña ── */
.modal-overlay {{
  display:none; position:fixed; inset:0;
  background:rgba(10,20,50,.55); backdrop-filter:blur(4px);
  align-items:center; justify-content:center; z-index:999;
}}
.modal-overlay.open {{ display:flex; }}
.modal-box {{
  background:#fff; border-radius:18px;
  box-shadow:0 24px 64px rgba(0,0,0,.35);
  padding:36px 32px 28px; width:90%; max-width:360px;
  animation:slideIn .3s ease-out; text-align:center;
}}
.modal-box h3 {{ color:#1a3a6b; font-size:17px; margin-bottom:6px; }}
.modal-box p  {{ color:#7f8c8d; font-size:13px; margin-bottom:20px; }}
.modal-box input {{
  width:100%; padding:13px 14px; border:2px solid #ecf0f1;
  border-radius:10px; font-size:16px; outline:none; transition:.25s;
  font-family:inherit; text-align:center; letter-spacing:3px;
  color:#1a3a6b; font-weight:700;
}}
.modal-box input:focus {{ border-color:#1a3a6b; box-shadow:0 0 0 3px rgba(26,58,107,.1); }}
.modal-box input.error {{ border-color:#e53935; animation:shake .3s; }}
@keyframes shake {{
  0%,100%{{transform:translateX(0)}} 25%{{transform:translateX(-6px)}} 75%{{transform:translateX(6px)}}
}}
.modal-actions {{ display:flex; gap:10px; margin-top:16px; }}
.modal-btn-ok {{
  flex:1; padding:13px; background:linear-gradient(135deg,#1a3a6b,#2d6ea8);
  color:#fff; border:none; border-radius:10px; font-size:14px;
  font-weight:700; cursor:pointer; transition:.2s;
}}
.modal-btn-ok:hover {{ transform:translateY(-2px); box-shadow:0 6px 18px rgba(26,58,107,.3); }}
.modal-btn-cancel {{
  flex:1; padding:13px; background:#f0f2f8;
  color:#7f8c8d; border:none; border-radius:10px; font-size:14px;
  font-weight:700; cursor:pointer; transition:.2s;
}}
.modal-btn-cancel:hover {{ background:#e2e6f0; }}
.modal-error {{ color:#e53935; font-size:12px; margin-top:8px; min-height:18px; }}
</style>
</head>
<body>

<div class="container">
  <div class="header">
    <div class="logo">&#127979;</div>
    <h1>{COLEGIO}</h1>
    <h2>{EXAMEN} &mdash; {ANIO}</h2>
  </div>

  <div class="tabs">
    <button class="tab-btn active" onclick="showTab('buscar',this)">&#128269;<br>Estudiantes</button>
    <button class="tab-btn" id="btn-informes" onclick="pedirClave(this)">&#128203;<br>Informes</button>
    <button class="tab-btn" onclick="showTab('descargas',this)">&#128190;<br>Descargar Notas</button>
  </div>

  <!-- ── BUSCAR ESTUDIANTE ── -->
  <div id="tab-buscar" class="section active">
    <div class="info-box">
      &#128161; Selecciona el grado e ingresa el nombre para consultar el informe.
    </div>
    <div class="field">
      <label class="field-label" for="sel-grado">Grado</label>
      <select id="sel-grado">
        <option value="">Todos los grados</option>
        {opciones_buscar}
      </select>
    </div>
    <div class="field">
      <label class="field-label" for="inp-nombre">Nombre o apellido</label>
      <input type="text" id="inp-nombre" placeholder="Escribe al menos 2 letras..." autocomplete="off"
             onkeydown="if(event.key==='Enter')buscar()">
    </div>
    <button class="search-btn" onclick="buscar()">&#128269; Buscar Informe</button>
    <div class="error-box" id="error-box"></div>
    <div id="resultados"></div>
  </div>

  <!-- ── INFORMES ── -->
  <div id="tab-docentes" class="section">
    <div class="subtabs">
      <button class="subtab-btn active" onclick="showSubtab('sub-directores',this)">&#128101; Director de Grupo</button>
      <button class="subtab-btn" onclick="showSubtab('sub-docentes',this)">&#128218; Informe Docente</button>
      <button class="subtab-btn" onclick="showSubtab('sub-desglose',this)">&#128202; Desglose por Pregunta</button>
    </div>
    <div id="sub-directores" class="subsection active">
      {rep_grupos}
    </div>
    <div id="sub-docentes" class="subsection">
      {rep_docentes}
    </div>
    <div id="sub-desglose" class="subsection">
      {rep_desglose}
    </div>
  </div>

  <!-- ── DESCARGAR NOTAS ── -->
  <div id="tab-descargas" class="section">
    <div class="info-box">
      &#128190; Selecciona el grado, el grupo y la asignatura para descargar las notas en CSV.
    </div>
    <div class="field">
      <label class="field-label">Grado</label>
      <select id="dl-grado" onchange="actualizarGrupos()">
        <option value="">-- Seleccione grado --</option>
        {opciones_dl}
      </select>
    </div>
    <div class="field">
      <label class="field-label">Grupo</label>
      <select id="dl-grupo">
        <option value="">-- Seleccione grupo --</option>
      </select>
    </div>
    <div class="field">
      <label class="field-label">Asignatura</label>
      <select id="dl-area">
        <option value="">-- Seleccione asignatura --</option>
      </select>
    </div>
    <button class="dl-btn" onclick="descargarCSV()">&#11015; Descargar CSV</button>
  </div>
</div>

<!-- ── Modal contraseña Informes ── -->
<div class="modal-overlay" id="modal-clave">
  <div class="modal-box">
    <h3>&#128274; Acceso restringido</h3>
    <p>Ingresa la contraseña para ver los informes</p>
    <input type="password" id="modal-input" placeholder="Contraseña"
           onkeydown="if(event.key==='Enter')verificarClave()">
    <div class="modal-error" id="modal-error"></div>
    <div class="modal-actions">
      <button class="modal-btn-cancel" onclick="cerrarModal()">Cancelar</button>
      <button class="modal-btn-ok" onclick="verificarClave()">&#128275; Ingresar</button>
    </div>
  </div>
</div>

<footer>
  {COLEGIO} &bull; {EXAMEN} {ANIO}
  <strong>SOFA EDITORES</strong>
</footer>

<script src="data.js"></script>
<script>
const AREAS_POR_GRADO = {json.dumps(areas_por_grado_js, ensure_ascii=False)};
var _informesDesbloqueado = false;
var _btnInformes = null;
function pedirClave(btn) {{
  if (_informesDesbloqueado) {{
    showTab('docentes', btn);
    return;
  }}
  _btnInformes = btn;
  document.getElementById('modal-input').value = '';
  document.getElementById('modal-error').textContent = '';
  document.getElementById('modal-input').classList.remove('error');
  document.getElementById('modal-clave').classList.add('open');
  setTimeout(function(){{ document.getElementById('modal-input').focus(); }}, 80);
}}
function verificarClave() {{
  var val = document.getElementById('modal-input').value;
  if (val === 'CESUM2026') {{
    _informesDesbloqueado = true;
    document.getElementById('modal-clave').classList.remove('open');
    showTab('docentes', _btnInformes);
  }} else {{
    var inp = document.getElementById('modal-input');
    inp.classList.add('error');
    document.getElementById('modal-error').textContent = 'Contraseña incorrecta. Inténtalo de nuevo.';
    inp.value = '';
    setTimeout(function(){{ inp.classList.remove('error'); }}, 400);
  }}
}}
function cerrarModal() {{
  document.getElementById('modal-clave').classList.remove('open');
}}
function showSubtab(id, btn) {{
  document.querySelectorAll('.subsection').forEach(function(s){{ s.classList.remove('active'); }});
  document.querySelectorAll('.subtab-btn').forEach(function(b){{ b.classList.remove('active'); }});
  document.getElementById(id).classList.add('active');
  if (btn) btn.classList.add('active');
}}
function showTab(id, btn) {{
  document.querySelectorAll('.section').forEach(function(s){{ s.classList.remove('active'); }});
  document.querySelectorAll('.tab-btn').forEach(function(b){{ b.classList.remove('active'); }});
  document.getElementById('tab-' + id).classList.add('active');
  if (btn) btn.classList.add('active');
  document.getElementById('resultados').innerHTML = '';
  document.getElementById('error-box').style.display = 'none';
}}

function norm(s) {{
  return String(s).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9 ]/g, ' ').trim();
}}
function badgeN(n) {{
  if (n >= 90) return '<span class="badge b-s">Superior</span>';
  if (n >= 70) return '<span class="badge b-a">Alto</span>';
  if (n >= 60) return '<span class="badge b-b">Basico</span>';
  return '<span class="badge b-l">Bajo</span>';
}}

function buscar() {{
  var grado  = document.getElementById('sel-grado').value;
  var query  = norm(document.getElementById('inp-nombre').value);
  var divRes = document.getElementById('resultados');
  var errBox = document.getElementById('error-box');
  errBox.style.display = 'none';
  divRes.innerHTML = '';

  if (!grado && query.length < 2) {{
    errBox.textContent = 'Selecciona un grado o escribe al menos 2 letras del nombre.';
    errBox.style.display = 'block';
    return;
  }}

  var res = ESTUDIANTES.filter(function(e) {{
    var okGrado  = !grado  || e.grupo.indexOf(grado) === 0;
    var okNombre = !query  || norm(e.nombre).indexOf(query) !== -1;
    return okGrado && okNombre;
  }});

  if (res.length === 0) {{
    divRes.innerHTML = '<div class="no-result">No se encontraron estudiantes.</div>';
    return;
  }}

  res.sort(function(a,b){{ return b.promedio - a.promedio; }});
  var shown = res.slice(0, 60);

  var html = shown.map(function(e) {{
    return '<a class="result-card" href="' + e.archivo + '" target="_blank">'
      + '<div>'
      + '<div class="r-nombre">' + e.nombre + '</div>'
      + '<div class="r-info">' + e.grupo + ' &bull; ID: ' + e.id + '</div>'
      + '</div>'
      + '<div class="r-nota">' + e.promedio.toFixed(1) + '<small>' + badgeN(e.promedio) + '</small></div>'
      + '</a>';
  }}).join('');

  if (res.length > 60) {{
    html += '<div class="no-result">Mostrando 60 de ' + res.length + '. Refina la busqueda.</div>';
  }}
  divRes.innerHTML = html;
}}

function actualizarGrupos() {{
  var grado = document.getElementById('dl-grado').value;
  var selG  = document.getElementById('dl-grupo');
  var selA  = document.getElementById('dl-area');
  selG.innerHTML = '<option value="">-- Seleccione grupo --</option>';
  selA.innerHTML = '<option value="">-- Seleccione asignatura --</option>';
  if (grado && GRUPOS_POR_GRADO[grado]) {{
    GRUPOS_POR_GRADO[grado].forEach(function(g) {{
      var opt = document.createElement('option');
      opt.value = g; opt.textContent = g;
      selG.appendChild(opt);
    }});
  }}
  if (grado && AREAS_POR_GRADO[grado]) {{
    AREAS_POR_GRADO[grado].forEach(function(a) {{
      var opt = document.createElement('option');
      opt.value = a; opt.textContent = a;
      selA.appendChild(opt);
    }});
  }}
}}

function descargarCSV() {{
  var grado = document.getElementById('dl-grado').value;
  var grupo = document.getElementById('dl-grupo').value;
  var area  = document.getElementById('dl-area').value;

  if (!grado || !grupo || !area) {{
    alert('Por favor selecciona grado, grupo y asignatura.');
    return;
  }}

  var datos = ESTUDIANTES.filter(function(e){{ return e.grupo === grupo; }});
  if (!datos.length) {{ alert('No hay estudiantes en ese grupo.'); return; }}

  datos.sort(function(a,b){{ return a.nombre.localeCompare(b.nombre); }});

  var lines = ['Codigo,Apellidos y Nombres,Grupo,' + area + ' (/100)'];
  datos.forEach(function(e) {{
    var nota = e.notas[area] !== undefined ? e.notas[area].toFixed(1) : '';
    lines.push('"' + e.id + '","' + e.nombre + '","' + e.grupo + '",' + nota);
  }});

  var csv  = '\\uFEFF' + lines.join('\\n');
  var blob = new Blob([csv], {{type:'text/csv;charset=utf-8;'}});
  var url  = URL.createObjectURL(blob);
  var a    = document.createElement('a');
  a.href = url;
  a.download = 'Notas_' + grupo.replace(/ /g,'_') + '_' + area.replace(/ /g,'_') + '.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}}
</script>
</body>
</html>"""

    (PORTAL_DIR / "index.html").write_text(portal, encoding='utf-8')
    print(f"\n  ✓ Portal generado: {PORTAL_DIR / 'index.html'}")

# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    PORTAL_DIR.mkdir(parents=True, exist_ok=True)

    SIP = "Sexto-Septimo IP"   # subcarpeta de sexto y septimo
    OD  = "Octavo- Decimo"     # subcarpeta de octavo, noveno y decimo

    print("\n📖 Leyendo tablas de respuestas...")
    tabla_sexto   = leer_tabla(BASE_DIR / SIP / "TABLA SEXTO.xlsx")
    tabla_septimo = leer_tabla(BASE_DIR / SIP / "TABLA SEPTIMO.xlsx")
    tabla_octavo  = leer_tabla(BASE_DIR / OD  / "tabla octavo.xlsx")
    tabla_noveno  = leer_tabla(BASE_DIR / OD  / "tabla noveno.xlsx")
    tabla_decimo  = leer_tabla(BASE_DIR / OD  / "Tabla decimo.xlsx")

    print("\n📊 Cargando datos de SEXTO...")
    s1_sexto = pd.concat([
        pd.read_csv(BASE_DIR / SIP / "SEXTO_Sesión1.csv"),
        pd.read_csv(BASE_DIR / "sexto01.csv"),
    ], ignore_index=True)
    s2_sexto = pd.concat([
        pd.read_csv(BASE_DIR / SIP / "SEXTO_Sesion2.csv"),
        pd.read_csv(BASE_DIR / "sextosesion2.csv"),
    ], ignore_index=True)
    df_sexto = cargar_grado(s1_sexto, s2_sexto)
    print(f"   {len(df_sexto)} estudiantes | grupos: {sorted(df_sexto['Grupo'].unique())}")

    print("\n📊 Cargando datos de SEPTIMO...")
    df_septimo = cargar_grado(f"{SIP}/SEPTIMO_Sesion 1.csv", f"{SIP}/SEPTIMO_Sesion 2.csv")
    print(f"   {len(df_septimo)} estudiantes | grupos: {sorted(df_septimo['Grupo'].unique())}")

    print("\n📊 Cargando datos de OCTAVO...")
    df_octavo = cargar_grado(
        f"{OD}/Octavo-1-all-Quiz Format 2026-03-20 11_03-2026-03-20 16_04_26.csv",
        f"{OD}/Octavo-2-all-Quiz Format 2026-03-20 11_05-2026-03-20 16_06_08.csv",
    )
    print(f"   {len(df_octavo)} estudiantes | grupos: {sorted(df_octavo['Grupo'].unique())}")

    print("\n📊 Cargando datos de NOVENO...")
    df_noveno = cargar_grado(
        f"{OD}/Noveno-1-all-Quiz Format 2026-03-20 11_07-2026-03-20 16_08_06.csv",
        f"{OD}/Noveno-2-all-Quiz Format 2026-03-20 11_08-2026-03-20 16_09_13.csv",
    )
    print(f"   {len(df_noveno)} estudiantes | grupos: {sorted(df_noveno['Grupo'].unique())}")

    print("\n📊 Cargando datos de DECIMO...")
    df_decimo = cargar_grado(
        f"{OD}/Decimo S1 IP 2026-all-Nombre-Cod-Pts-2026-03-20 16_44_42.csv",
        f"{OD}/Decimo S2 IP 2026-all-Nombre-Cod-Pts-2026-03-20 16_44_57.csv",
        areas_preguntas=AREAS_PREGUNTAS_DECIMO,
    )
    print(f"   {len(df_decimo)} estudiantes | grupos: {sorted(df_decimo['Grupo'].unique())}")

    todos_individuales = []
    todos_docentes     = []
    todos_directores   = []
    todos_desgloses    = []

    COLORES_DECIMO = {**COLORES_AREA, **{"Filosofia":"#9B59B6","C. Politicas":"#1ABC9C","Quimica":"#E74C3C","Fisica":"#F39C12"}}

    for df, gstr, tabla, ap, ca in [
        (df_sexto,   "Sexto",   tabla_sexto,   AREAS_PREGUNTAS,       COLORES_AREA),
        (df_septimo, "Septimo", tabla_septimo, AREAS_PREGUNTAS,       COLORES_AREA),
        (df_octavo,  "Octavo",  tabla_octavo,  AREAS_PREGUNTAS,       COLORES_AREA),
        (df_noveno,  "Noveno",  tabla_noveno,  AREAS_PREGUNTAS,       COLORES_AREA),
        (df_decimo,  "Decimo",  tabla_decimo,  AREAS_PREGUNTAS_DECIMO, COLORES_DECIMO),
    ]:
        print(f"\n── Informes individuales ({gstr}) ──")
        todos_individuales += informes_individuales(df, gstr, tabla, areas_preguntas=ap)

        print(f"\n── Informes docentes ({gstr}) ──")
        todos_docentes += informes_docentes(df, gstr, tabla, areas_preguntas=ap, colores_area=ca)

        print(f"\n── Informes directores ({gstr}) ──")
        todos_directores += informes_directores(df, gstr, areas_preguntas=ap, colores_area=ca)

        print(f"\n── Desglose por pregunta ({gstr}) ──")
        todos_desgloses += desglose_preguntas(df, gstr, tabla, areas_preguntas=ap, colores_area=ca)

    print("\n── Generando portal web ──")
    generar_portal(
        {"Sexto": df_sexto, "Septimo": df_septimo,
         "Octavo": df_octavo, "Noveno": df_noveno, "Decimo": df_decimo},
        todos_docentes,
        todos_directores,
        archivos_desglose=todos_desgloses,
        areas_por_grado={
            "Sexto":   list(AREAS_PREGUNTAS.keys()),
            "Septimo": list(AREAS_PREGUNTAS.keys()),
            "Octavo":  list(AREAS_PREGUNTAS.keys()),
            "Noveno":  list(AREAS_PREGUNTAS.keys()),
            "Decimo":  list(AREAS_PREGUNTAS_DECIMO.keys()),
        }
    )

    print(f"\n✅ Todo generado en: {PORTAL_DIR}")
    print(f"   Individuales: {len(todos_individuales)} archivos")
    print(f"   Docentes:     {len(todos_docentes)} archivos")
    print(f"   Directores:   {len(todos_directores)} archivos")

if __name__ == "__main__":
    main()
