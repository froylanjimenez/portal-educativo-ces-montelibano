#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador de informes — SEGUNDO PERIODO
Concentracion Educativa del Sur de Montelibano

Trabaja a partir de data.js (notas por area de Primer Periodo) y genera
resultados similares para el Segundo Periodo, regenerando:
  - Informes individuales
  - Informes docentes (por area)        [sin analisis por pregunta]
  - Informes de directores de grupo
NO genera la hoja de respuestas (desglose por pregunta correcta/incorrecta).

Reproducible: usa una semilla fija para que las notas no cambien entre corridas.
"""

import os, re, json, math, random, unicodedata, statistics
from pathlib import Path

# ─── Config ───────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).resolve().parent
PORTAL_DIR = BASE_DIR                      # los informes viven junto al portal
BASE_JS    = BASE_DIR / "data_periodo1.js" # base inmutable (notas reales del Primer Periodo)
DATA_JS    = BASE_DIR / "data.js"          # salida que consume el portal (Segundo Periodo)

COLEGIO = "Concentracion Educativa del Sur de Montelibano"
EXAMEN  = "Examen Final de Segundo Periodo"
ANIO    = "2026"
PERIODO = "Segundo Periodo"

SEED = 20262           # semilla fija -> resultados estables

# Overrides manuales de notas de Segundo Periodo: (id, area) -> nota forzada.
# Se aplican DESPUES de la generacion aleatoria (no alteran las demas notas) y
# respetan el redondeo al puntaje posible segun preguntas por area.
OVERRIDES = {
    ("802008", "Matematicas"): 15.0,  # Felipe Beltran Diaz: correccion (0 -> similar al Primer Periodo)
}

# Estudiantes agregados manualmente al Segundo Periodo (sin Primer Periodo previo).
# Se inyectan DESPUES de la generacion aleatoria (no alteran el stream ni las notas
# de los demas) y reciben notas cercanas al promedio de su grupo, respetando el
# puntaje posible segun preguntas por area.
NUEVOS_ESTUDIANTES = [
    {"id": "903039", "nombre": "EILEEN GONZALEZ POLO", "grupo": "Noveno C", "grado": "Noveno",
     "archivo": "informes/Noveno/individuales/903039_GONZALEZ_POLO_EILEEN.html"},
    {"id": "903040", "nombre": "LINDA RESTAN MARTINEZ", "grupo": "Noveno C", "grado": "Noveno",
     "archivo": "informes/Noveno/individuales/903040_RESTAN_MARTINEZ_LINDA.html"},
    {"id": "101033", "nombre": "JESUS VERGARA", "grupo": "Decimo A", "grado": "Decimo",
     "archivo": "informes/Decimo/individuales/101033_VERGARA_JESUS.html"},
    {"id": "604036", "nombre": "MARIA HOYOS", "grupo": "Sexto D", "grado": "Sexto",
     "archivo": "informes/Sexto/individuales/604036_HOYOS_MARIA.html"},
    {"id": "604037", "nombre": "SANTIAGO BARRETO", "grupo": "Sexto D", "grado": "Sexto",
     "archivo": "informes/Sexto/individuales/604037_BARRETO_SANTIAGO.html"},
    {"id": "605031", "nombre": "SAILETH AVILA", "grupo": "Sexto E", "grado": "Sexto",
     "archivo": "informes/Sexto/individuales/605031_AVILA_SAILETH.html"},
    {"id": "605032", "nombre": "DANIEL BERRIO", "grupo": "Sexto E", "grado": "Sexto",
     "archivo": "informes/Sexto/individuales/605032_BERRIO_DANIEL.html"},
    {"id": "905035", "nombre": "ANDRES DIAZ MEZA", "grupo": "Noveno E", "grado": "Noveno",
     "archivo": "informes/Noveno/individuales/905035_DIAZ_MEZA_ANDRES.html"},
    {"id": "703038", "nombre": "MARYSOL GONZALEZ DIAZ", "grupo": "Septimo C", "grado": "Septimo",
     "archivo": "informes/Septimo/individuales/703038_GONZALEZ_DIAZ_MARYSOL.html"},
]

# Re-siembra de Religion para grupos cuya clave de respuestas salio anomala y
# dejo a todo el grupo perdiendo (grados 6-9, cuadernillo con clave inconsistente
# en Religion). Para cada estudiante del grupo se re-siembra SOLO la nota de
# Religion, centrada en el objetivo del grupo con variacion leve y reproducible
# (semilla por id, independiente del stream aleatorio general -> no altera las
# demas notas ni las de otros estudiantes), respetando el puntaje posible por
# area (multiplos de 100/n_preg). Se aplica ANTES de OVERRIDES, de modo que un
# override manual por estudiante sigue teniendo prioridad.
AJUSTE_RELIGION = {
    "Septimo A": 78.0, "Septimo B": 78.0, "Septimo C": 78.0,
    "Septimo D": 78.0, "Septimo E": 78.0, "Septimo F": 78.0,
}
RELIGION_SPREAD  = 7.0    # desviacion de la variacion por estudiante
RELIGION_MINIMO  = 57.1   # piso antes de redondear (evita colas bajas)

# Estudiantes retirados (no presentaron la prueba): se filtran DESPUES de la
# generacion aleatoria para no alterar el stream ni las notas de los demas.
REMOVIDOS = {
    "605010",  # Gabriela Maria Hernandez Sierra (Sexto E): no estuvo en la prueba
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
    "Filosofia":          "#9B59B6",
    "C. Politicas":       "#1ABC9C",
    "Quimica":            "#E74C3C",
    "Fisica":             "#F39C12",
}

GRADE_NUM = {"Sexto":"6","Septimo":"7","Octavo":"8","Noveno":"9","Decimo":"10"}

# ─── Numero de preguntas por area (segun cuadernillos de la evaluacion) ────────
# Cada pregunta vale 100/N, por lo que las notas posibles son multiplos de ese
# valor. Los puntajes se redondean HACIA ARRIBA al valor posible mas cercano.
PREGUNTAS_6A9 = {   # Sexto, Septimo, Octavo y Noveno (cuadernillo grados 6°-9°)
    "Matematicas": 20, "Etica": 7,  "Lengua Castellana": 20, "Ed. Fisica": 7,
    "Ciencias Sociales": 20, "Gestion": 7, "Ciencias Naturales": 20,
    "Artistica": 7, "Ingles": 20, "Religion": 7, "Tecnologia": 7,
}
PREGUNTAS_10 = {    # Decimo (cuadernillo grado 10°)
    "Matematicas": 20, "Filosofia": 6, "Lengua Castellana": 20, "Etica": 6,
    "Ciencias Sociales": 20, "C. Politicas": 6, "Quimica": 10, "Ed. Fisica": 6,
    "Fisica": 10, "Religion": 6, "Ingles": 20, "Tecnologia": 6,
}

def preguntas_de(grado):
    return PREGUNTAS_10 if grado == "Decimo" else PREGUNTAS_6A9

def redondear_posible(valor, n_preg):
    """Redondea 'valor' (0-100) HACIA ARRIBA al puntaje posible mas cercano
    para un area de n_preg preguntas (cada pregunta vale 100/n_preg).
    Ej.: 6 preguntas -> paso 16.67; un 20 sube al siguiente posible: 33.3."""
    if not n_preg:
        return valor
    paso = 100.0 / n_preg
    k = math.ceil(valor / paso - 1e-9)        # nro. de respuestas correctas (hacia arriba)
    k = max(0, min(n_preg, k))
    return round(k * paso, 1)

def redondear_cercano(valor, n_preg):
    """Redondea 'valor' al puntaje posible MAS CERCANO (multiplo de 100/n_preg).
    Se usa al sembrar notas de estudiantes nuevos centradas en un objetivo
    (p.ej. el promedio del grupo), evitando el sesgo al alza del redondeo hacia
    arriba. El resultado sigue siendo un puntaje alcanzable segun las preguntas."""
    if not n_preg:
        return valor
    paso = 100.0 / n_preg
    k = max(0, min(n_preg, round(valor / paso)))
    return round(k * paso, 1)

# ─── Helpers de presentacion (mismos estilos del generador original) ───────────

def sanitize(s):
    s = ''.join(c for c in unicodedata.normalize('NFD', str(s))
                if unicodedata.category(c) != 'Mn')
    return re.sub(r'[^A-Za-z0-9_\-]', '_', s)

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

# ─── Carga de datos (data.js) y generacion del Segundo Periodo ─────────────────

def cargar_data_js():
    # Siempre se parte de la base inmutable del Primer Periodo (idempotente).
    fuente = BASE_JS if BASE_JS.exists() else DATA_JS
    txt = fuente.read_text(encoding='utf-8')
    m  = re.search(r'const ESTUDIANTES\s*=\s*(\[.*?\]);', txt, re.S)
    m2 = re.search(r'const GRUPOS_POR_GRADO\s*=\s*(\{.*?\});', txt, re.S)
    estudiantes = json.loads(m.group(1))
    grupos      = json.loads(m2.group(1))
    return estudiantes, grupos

def perturbar(valor, rnd):
    """Genera una nota de segundo periodo 'similar' a la del primero."""
    nv = valor + rnd.gauss(0, 6.5)        # variacion moderada, centrada en 0
    return round(min(100.0, max(0.0, nv)), 1)

def generar_segundo_periodo(estudiantes):
    rnd = random.Random(SEED)
    nuevos = []
    for e in estudiantes:
        areas = list(e["notas"].keys())          # conserva el orden canonico
        preg  = preguntas_de(e["grado"])          # nro. de preguntas por area del grado
        notas = {a: redondear_posible(perturbar(e["notas"][a], rnd), preg.get(a))
                 for a in areas}
        obj = AJUSTE_RELIGION.get(e["grupo"])             # re-siembra de Religion (grupo con clave anomala)
        if obj is not None and "Religion" in notas:
            rr  = random.Random(int(str(e["id"])) * 31 + 7)   # reproducible e indep. del stream general
            val = min(100.0, max(RELIGION_MINIMO, obj + rr.gauss(0, RELIGION_SPREAD)))
            notas["Religion"] = redondear_cercano(val, preg.get("Religion"))
        for a in areas:                                   # overrides manuales (post-aleatorio)
            ov = OVERRIDES.get((str(e["id"]), a))
            if ov is not None:
                notas[a] = redondear_posible(ov, preg.get(a))
        prom_full = statistics.fmean(notas.values())
        nuevos.append({
            "id":       str(e["id"]),
            "nombre":   e["nombre"],
            "grupo":    e["grupo"],
            "grado":    e["grado"],
            "areas":    areas,
            "notas":    notas,
            "prom":     prom_full,
            "promedio": round(prom_full, 1),
            "archivo":  e["archivo"],
        })
    return nuevos

def build_nuevos(alumnos):
    """Construye los estudiantes agregados manualmente (NUEVOS_ESTUDIANTES) con
    notas cercanas al promedio de su grupo (Segundo Periodo), con variacion leve
    y reproducible, respetando el puntaje posible por area."""
    por_grupo = {}
    for a in alumnos:
        por_grupo.setdefault(a["grupo"], []).append(a)
    nuevos = []
    for meta in NUEVOS_ESTUDIANTES:
        miembros = por_grupo[meta["grupo"]]
        areas = miembros[0]["areas"]                 # orden canonico del grado
        preg  = preguntas_de(meta["grado"])
        rnd   = random.Random(int(meta["id"]))       # variacion leve por estudiante
        notas = {}
        for ar in areas:
            prom_area = statistics.fmean(m["notas"][ar] for m in miembros)
            notas[ar] = redondear_cercano(prom_area + rnd.gauss(0, 4.0), preg.get(ar))
        prom_full = statistics.fmean(notas.values())
        nuevos.append({
            "id":       meta["id"],
            "nombre":   meta["nombre"],
            "grupo":    meta["grupo"],
            "grado":    meta["grado"],
            "areas":    areas,
            "notas":    notas,
            "prom":     prom_full,
            "promedio": round(prom_full, 1),
            "archivo":  meta["archivo"],
        })
    return nuevos

def rank_min(valores, idx):
    """Posicion 1..N (mayor=1), empates con minimo, igual a pandas method='min'."""
    v = valores[idx]
    return 1 + sum(1 for x in valores if x > v)

# ─── 1. Informes individuales ──────────────────────────────────────────────────

def informes_individuales(alumnos_grado, grado_str):
    proms = [a["prom"] for a in alumnos_grado]
    total_grado = len(alumnos_grado)
    # tam de grupo
    tam_grupo = {}
    for a in alumnos_grado:
        tam_grupo[a["grupo"]] = tam_grupo.get(a["grupo"], 0) + 1

    for i, a in enumerate(alumnos_grado):
        prom = a["prom"]
        nv, _ = nivel_desempeno(prom)
        pg_grado = rank_min(proms, i)
        # pos en grupo
        proms_grupo = [x["prom"] for x in alumnos_grado if x["grupo"] == a["grupo"]]
        pg = 1 + sum(1 for x in proms_grupo if x > prom)
        tg_g = tam_grupo[a["grupo"]]

        rows_notas = ""
        recs = ""
        for area in a["areas"]:
            nota = a["notas"][area]
            nva, _ = nivel_desempeno(nota)
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
<title>Informe - {a['nombre']}</title>
{css()}
</head>
<body>
<div class="page">
  {encabezado(f'Informe Individual &mdash; {a["grupo"]}')}
  <div class="grid2">
    <div><div class="info-label">Estudiante</div><div class="info-val">{a['nombre']}</div></div>
    <div><div class="info-label">Grupo</div><div class="info-val">{a['grupo']}</div></div>
    <div><div class="info-label">Codigo</div><div class="info-val">{a['id']}</div></div>
    <div><div class="info-label">Periodo</div><div class="info-val">{PERIODO} {ANIO}</div></div>
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
      <div class="slabel">{a['grupo']}</div>
    </div>
    <div style="text-align:center;">
      <div class="slabel">Posicion en grado</div>
      <div class="snum">{pg_grado}<span style="font-size:14px;">/{total_grado}</span></div>
      <div class="slabel">{a['grupo'].split()[0]}</div>
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

        fpath = PORTAL_DIR / a["archivo"]
        fpath.parent.mkdir(parents=True, exist_ok=True)
        fpath.write_text(html, encoding='utf-8')

    print(f"  ✓ {total_grado} informes individuales — {grado_str}")

# ─── 2. Informes docentes (por area, sin analisis por pregunta) ────────────────

def informes_docentes(alumnos_grado, grado_str, areas, colores_area):
    out = PORTAL_DIR / "informes" / grado_str / "docentes"
    out.mkdir(parents=True, exist_ok=True)
    grupos = sorted({a["grupo"] for a in alumnos_grado})
    total  = len(alumnos_grado)
    archivos = []

    for area in areas:
        notas_area = [a["notas"][area] for a in alumnos_grado]
        prom_g = statistics.fmean(notas_area)
        bajo     = sum(1 for n in notas_area if n < 60)
        basico   = sum(1 for n in notas_area if 60 <= n < 70)
        alto     = sum(1 for n in notas_area if 70 <= n < 90)
        superior = sum(1 for n in notas_area if n >= 90)
        aprob    = sum(1 for n in notas_area if n >= 60)

        rows_grupos = ""
        for g in grupos:
            sub = [a for a in alumnos_grado if a["grupo"] == g]
            ng = len(sub)
            pm = statistics.fmean(a["notas"][area] for a in sub)
            apro = sum(1 for a in sub if a["notas"][area] >= 60)
            nv, _ = nivel_desempeno(pm)
            rows_grupos += f"""
            <tr>
              <td><strong>{g}</strong></td>
              <td style="text-align:center;">{ng}</td>
              <td>{barra(pm, colores_area.get(area,'#555'))}</td>
              <td style="text-align:center;font-weight:700;">{pm:.1f}</td>
              <td style="text-align:center;">{badge(nv)}</td>
              <td style="text-align:center;">{apro} ({apro/ng*100:.0f}%)</td>
            </tr>"""

        ordenados = sorted(alumnos_grado, key=lambda a: a["notas"][area], reverse=True)
        top5, bot5 = ordenados[:5], ordenados[-5:][::-1]
        r_top = "".join(f"<tr><td>{a['nombre']}</td><td>{a['grupo']}</td>"
                        f"<td style='text-align:center;font-weight:700;color:#1a7a1a;'>{a['notas'][area]:.1f}</td></tr>"
                        for a in top5)
        r_bot = "".join(f"<tr><td>{a['nombre']}</td><td>{a['grupo']}</td>"
                        f"<td style='text-align:center;font-weight:700;color:#cc2200;'>{a['notas'][area]:.1f}</td></tr>"
                        for a in bot5)

        pages = [f"""
        <div class="page">
          {encabezado(f'Informe Docente &mdash; {area} &mdash; {grado_str}')}
          <div class="sumbox">
            <div><div class="slabel">PROMEDIO GRADO</div>
                 <div class="snum">{prom_g:.1f}<span style="font-size:16px;">/100</span></div></div>
            <div style="text-align:center;"><div class="slabel">Evaluados</div>
                 <div class="snum">{total}</div></div>
            <div style="text-align:center;"><div class="slabel">Aprobados (&#8805;60)</div>
                 <div class="snum">{aprob} <span style="font-size:14px;">({aprob/total*100:.0f}%)</span></div></div>
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
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:12px;">
            <div><div class="sec" style="margin-top:0;">&#127942; Mejores puntajes</div>
              <table><tr><th>Estudiante</th><th>Grupo</th>
                  <th style="text-align:center;">Nota</th></tr>{r_top}</table></div>
            <div><div class="sec" style="margin-top:0;">&#128204; Atencion prioritaria</div>
              <table><tr><th>Estudiante</th><th>Grupo</th>
                  <th style="text-align:center;">Nota</th></tr>{r_bot}</table></div>
          </div>
          {pie()}
        </div>"""]

        for g in grupos:
            sub = sorted((a for a in alumnos_grado if a["grupo"] == g),
                         key=lambda a: a["notas"][area], reverse=True)
            notas_g = [a["notas"][area] for a in sub]
            rows_est = ""
            for i, a in enumerate(sub):
                nva, _ = nivel_desempeno(a["notas"][area])
                rows_est += f"""
                <tr>
                  <td style="text-align:center;">{i+1}</td>
                  <td><strong>{a['nombre']}</strong></td>
                  <td>{barra(a['notas'][area], colores_area.get(area,'#555'))}</td>
                  <td style="text-align:center;font-weight:700;">{a['notas'][area]:.1f}</td>
                  <td style="text-align:center;">{badge(nva)}</td>
                </tr>"""
            pages.append(f"""
            <div class="page">
              {encabezado(f'Informe Docente &mdash; {area} &mdash; {g}')}
              <div class="grid2">
                <div><div class="info-label">Promedio del grupo</div>
                     <div class="info-val">{statistics.fmean(notas_g):.1f}</div></div>
                <div><div class="info-label">Estudiantes</div>
                     <div class="info-val">{len(sub)}</div></div>
                <div><div class="info-label">Nota maxima</div>
                     <div class="info-val">{max(notas_g):.1f}</div></div>
                <div><div class="info-label">Nota minima</div>
                     <div class="info-val">{min(notas_g):.1f}</div></div>
              </div>
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
        html  = f'<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Docente {area} {grado_str}</title>{css()}</head><body>'
        html += "\n".join(pages) + "</body></html>"
        (out / fname).write_text(html, encoding='utf-8')
        archivos.append({"area": area, "grado": grado_str})

    print(f"  ✓ {len(archivos)} informes docentes — {grado_str}")

# ─── 3. Informes directores de grupo ──────────────────────────────────────────

def informes_directores(alumnos_grado, grado_str, areas, colores_area):
    out = PORTAL_DIR / "informes" / grado_str / "directores"
    out.mkdir(parents=True, exist_ok=True)
    grupos = sorted({a["grupo"] for a in alumnos_grado})
    proms_grado = [a["prom"] for a in alumnos_grado]
    prom_grado  = statistics.fmean(proms_grado)
    total_grado = len(alumnos_grado)
    proms_area_grado = {ar: statistics.fmean(a["notas"][ar] for a in alumnos_grado) for ar in areas}
    pos_grado_map = {a["id"]: rank_min(proms_grado, i) for i, a in enumerate(alumnos_grado)}

    for grupo in grupos:
        sub = [a for a in alumnos_grado if a["grupo"] == grupo]
        sub = sorted(sub, key=lambda a: a["prom"], reverse=True)
        proms_sub = [a["prom"] for a in sub]
        n = len(sub)
        prom_g = statistics.fmean(proms_sub)
        diff   = prom_g - prom_grado
        arrow  = ('&#9650; por encima' if diff >= 0 else '&#9660; por debajo') + f' ({abs(diff):.1f} pts)'
        acolor = '#1a7a1a' if diff >= 0 else '#cc2200'

        bajo     = sum(1 for p in proms_sub if p < 60)
        basico   = sum(1 for p in proms_sub if 60 <= p < 70)
        alto     = sum(1 for p in proms_sub if 70 <= p < 90)
        superior = sum(1 for p in proms_sub if p >= 90)

        rows_est = ""
        for i, a in enumerate(sub):
            nv, _ = nivel_desempeno(a["prom"])
            pos_loc = 1 + sum(1 for x in proms_sub if x > a["prom"])
            pos_grado_est = pos_grado_map[a["id"]]
            rows_est += f"""
            <tr>
              <td style="text-align:center;">{pos_loc}</td>
              <td><strong>{a['nombre']}</strong></td>
              <td>{barra(a['prom'])}</td>
              <td style="text-align:center;font-weight:700;">{a['prom']:.1f}</td>
              <td style="text-align:center;">{badge(nv)}</td>
              <td style="text-align:center;">{pos_grado_est}/{total_grado}</td>
            </tr>"""

        rows_areas = ""
        prom_area_sub = {}
        for area in areas:
            pm_a = statistics.fmean(a["notas"][area] for a in sub)
            prom_area_sub[area] = pm_a
            pm_g = proms_area_grado[area]
            d = pm_a - pm_g
            ar = f'<span style="color:{"#1a7a1a" if d>=0 else "#cc2200"}">{"&#9650;" if d>=0 else "&#9660;"} {abs(d):.1f}</span>'
            nv, _ = nivel_desempeno(pm_a)
            rows_areas += f"""
            <tr>
              <td><strong>{area}</strong></td>
              <td>{barra(pm_a, colores_area.get(area,'#555'))}</td>
              <td style="text-align:center;font-weight:700;">{pm_a:.1f}</td>
              <td style="text-align:center;">{pm_g:.1f}</td>
              <td style="text-align:center;">{ar}</td>
              <td style="text-align:center;">{badge(nv)}</td>
            </tr>"""

        th_areas_desglose = "".join(
            f'<th style="text-align:center;background:{colores_area.get(area,"#555")};color:#fff;'
            f'font-size:10px;padding:5px 3px;min-width:55px;">{area}</th>'
            for area in areas
        )
        bg_nivel = {"Superior":"#e8f5e9","Alto":"#e3f0fa","Basico":"#fff8e1","Bajo":"#ffebee"}
        fg_nivel = {"Superior":"#1a7a1a","Alto":"#2d7fcc","Basico":"#d48a00","Bajo":"#cc2200"}

        rows_desglose = ""
        for i, a in enumerate(sub):
            pos_loc = 1 + sum(1 for x in proms_sub if x > a["prom"])
            celdas = ""
            for area in areas:
                na = a["notas"][area]
                nv_a, _ = nivel_desempeno(na)
                celdas += (f'<td style="text-align:center;background:{bg_nivel.get(nv_a,"#fff")};'
                           f'color:{fg_nivel.get(nv_a,"#333")};font-weight:700;font-size:12px;">{na:.1f}</td>')
            rows_desglose += f"""
            <tr>
              <td style="text-align:center;">{pos_loc}</td>
              <td style="white-space:nowrap;"><strong>{a['nombre']}</strong></td>
              <td style="text-align:center;font-weight:700;">{a['prom']:.1f}</td>
              {celdas}
            </tr>"""

        riesgo = sorted(
            [a for a in sub if sum(1 for ar in areas if a["notas"][ar] < 60) >= 3],
            key=lambda a: sum(1 for ar in areas if a["notas"][ar] < 60), reverse=True)
        if riesgo:
            rows_r = "".join(
                f"<tr><td><strong>{a['nombre']}</strong></td>"
                f"<td style='text-align:center;'>{sum(1 for ar in areas if a['notas'][ar]<60)}</td>"
                f"<td style='font-size:11px;'>{', '.join(ar for ar in areas if a['notas'][ar]<60)}</td>"
                f"<td style='text-align:center;font-weight:700;color:#cc2200;'>{a['prom']:.1f}</td></tr>"
                for a in riesgo)
            sec_riesgo = f"""
            <div class="alert-w">&#9888; <strong>{len(riesgo)} estudiante(s)</strong> con Bajo en 3 o mas areas. Seguimiento prioritario recomendado.</div>
            <table><tr><th>Estudiante</th><th style="text-align:center;">Areas en Bajo</th>
                <th>Areas criticas</th><th style="text-align:center;">Promedio</th></tr>{rows_r}</table>"""
        else:
            sec_riesgo = '<div class="alert-ok">&#10003; Ningun estudiante con Bajo en 3 o mas areas.</div>'

        mejor_a = max(prom_area_sub, key=prom_area_sub.get)
        peor_a  = min(prom_area_sub, key=prom_area_sub.get)
        aprob_grupo = sum(1 for p in proms_sub if p >= 60)

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
    <div><div class="info-label">Periodo</div><div class="info-val">{PERIODO} {ANIO}</div></div>
  </div>
  <div class="sumbox">
    <div><div class="slabel">PROMEDIO DEL GRUPO</div>
         <div class="snum">{prom_g:.1f}<span style="font-size:16px;">/100</span></div></div>
    <div style="text-align:center;"><div class="slabel">vs. promedio grado</div>
         <div class="snum">{prom_grado:.1f}</div>
         <div style="font-size:11px;color:{acolor};">{arrow}</div></div>
    <div style="text-align:center;"><div class="slabel">Aprobados (&#8805;60)</div>
         <div class="snum">{aprob_grupo}<span style="font-size:14px;">/{n}</span></div></div>
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
  <div class="alert-i">&#128170; Fortaleza: <strong>{mejor_a}</strong> ({prom_area_sub[mejor_a]:.1f}/100) &nbsp;&nbsp;
       &#128202; A reforzar: <strong>{peor_a}</strong> ({prom_area_sub[peor_a]:.1f}/100)</div>
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

        (out / f"Director_{sanitize(grupo)}.html").write_text(html, encoding='utf-8')

    print(f"  ✓ {len(grupos)} informes directores — {grado_str}")

# ─── data.js de salida ─────────────────────────────────────────────────────────

def escribir_data_js(alumnos, grupos_por_grado):
    # Posiciones por grado y por grupo (mismo criterio que los informes: empate=min)
    pos = {}
    por_grado = {}
    for a in alumnos:
        por_grado.setdefault(a["grado"], []).append(a)
    for lst in por_grado.values():
        proms = [x["prom"] for x in lst]
        total_grado = len(lst)
        por_grupo = {}
        for x in lst:
            por_grupo.setdefault(x["grupo"], []).append(x["prom"])
        for x in lst:
            pg_grado = 1 + sum(1 for p in proms if p > x["prom"])
            gp = por_grupo[x["grupo"]]
            pg_grupo = 1 + sum(1 for p in gp if p > x["prom"])
            pos[x["id"]] = (pg_grupo, len(gp), pg_grado, total_grado)

    estudiantes_json = [{
        "id":          a["id"],
        "nombre":      a["nombre"],
        "grupo":       a["grupo"],
        "grado":       a["grado"],
        "promedio":    a["promedio"],
        "notas":       {ar: round(a["notas"][ar], 1) for ar in a["areas"]},
        "pos_grupo":   pos[a["id"]][0],
        "total_grupo": pos[a["id"]][1],
        "pos_grado":   pos[a["id"]][2],
        "total_grado": pos[a["id"]][3],
        "archivo":     a["archivo"],
    } for a in alumnos]
    data_js = (f"const ESTUDIANTES = {json.dumps(estudiantes_json, ensure_ascii=False)};\n\n"
               f"const GRUPOS_POR_GRADO = {json.dumps(grupos_por_grado, ensure_ascii=False)};\n")
    DATA_JS.write_text(data_js, encoding='utf-8')
    print(f"  ✓ data.js actualizado ({len(estudiantes_json)} estudiantes)")

# ─── Main ──────────────────────────────────────────────────────────────────────

def main():
    print(f"📖 Leyendo base de Primer Periodo ({BASE_JS.name if BASE_JS.exists() else DATA_JS.name})...")
    estudiantes, grupos_por_grado = cargar_data_js()
    print(f"   {len(estudiantes)} estudiantes")

    print(f"\n🎲 Generando resultados similares — {PERIODO} (semilla {SEED})...")
    alumnos = generar_segundo_periodo(estudiantes)
    if REMOVIDOS:
        antes = len(alumnos)
        alumnos = [a for a in alumnos if a["id"] not in REMOVIDOS]
        print(f"   - {antes - len(alumnos)} estudiantes retirados")
    nuevos = build_nuevos(alumnos)                    # agregados manualmente (post-generacion)
    if nuevos:
        alumnos += nuevos
        print(f"   + {len(nuevos)} estudiantes agregados manualmente")

    # Agrupar por grado conservando el orden de aparicion
    grados = []
    por_grado = {}
    for a in alumnos:
        if a["grado"] not in por_grado:
            por_grado[a["grado"]] = []
            grados.append(a["grado"])
        por_grado[a["grado"]].append(a)

    for grado_str in grados:
        ag = por_grado[grado_str]
        areas = ag[0]["areas"]                      # orden canonico del grado
        print(f"\n── {grado_str} ({len(ag)} estudiantes) ──")
        informes_individuales(ag, grado_str)
        informes_docentes(ag, grado_str, areas, COLORES_AREA)
        informes_directores(ag, grado_str, areas, COLORES_AREA)

    print("\n── data.js ──")
    escribir_data_js(alumnos, grupos_por_grado)

    print(f"\n✅ Segundo Periodo generado en: {PORTAL_DIR}")

if __name__ == "__main__":
    main()
