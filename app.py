# -*- coding: utf-8 -*-
"""
Servidor del Asesor de Portafolios Inteligentes - modelo del TAF  ·  versión 3.0
Recibe los datos que envía el agente de Copilot Studio, llama al modelo y
devuelve el resultado en JSON. Desde la versión 1.1 también guarda el informe
que redacta el agente y entrega un enlace para descargarlo como archivo .md;
desde la versión 1.2 entrega además un segundo enlace con el informe en Word (.docx);
desde la versión 1.3 ese Word trae seis gráficos al estilo TradingView (graficos.py), y
desde la versión 1.31 cada línea lleva un comentario explicativo (el funcionamiento es idéntico al de la 1.3), y
desde la versión 3.0 la numeración se unifica con el material del curso (el funcionamiento es idéntico al de la 1.31).

Rutas:
  GET  /                  información del servicio
  GET  /health            comprobación de que el servidor está vivo
  POST /analizar          cálculo del portafolio
  POST /informe           guarda el informe del agente y devuelve el enlace de descarga
  GET  /descargar/{id}    descarga el informe como archivo .md
  GET  /descargar/{id}/word  descarga el mismo informe como archivo de Word (.docx)
  GET  /docs              página de prueba que genera FastAPI
"""
# Qué es una ruta: una dirección del servidor a la que se le puede pedir algo.
#   GET  = pedir información (como abrir una página web).
#   POST = enviar datos para que el servidor los procese (como enviar un formulario).
# Ejemplo: el agente envía {"monto": 100000, "perfil": "conservador", ...} por POST a /analizar.

import io                                      # io: archivos en memoria (para armar el Word sin tocar el disco).
import re                                      # re: expresiones regulares, patrones para buscar texto.
import unicodedata                             # unicodedata: permite quitar tildes ("Compañía" → "Compania").
import uuid                                    # uuid: genera identificadores únicos al azar para cada informe.
from datetime import datetime, timedelta, timezone   # Fechas, horas y zonas horarias.
from typing import Optional                    # Optional: indica que un dato puede venir o no.

from fastapi import FastAPI, HTTPException, Request  # FastAPI: el servidor; HTTPException: respuestas de aviso
                                                     # (400, 404); Request: datos de la petición recibida.
from fastapi.responses import Response         # Response: para devolver archivos (.md o .docx) en vez de JSON.
from pydantic import BaseModel, Field          # pydantic: define y valida la forma de los datos que llegan.
from docx import Document                      # python-docx: crea documentos de Word.
from docx.enum.text import WD_ALIGN_PARAGRAPH  # Alineaciones de párrafo (centrado, justificado, ...).
from docx.oxml import OxmlElement              # Acceso al XML interno del Word (para sombrear celdas).
from docx.oxml.ns import qn                    # qn: escribe los nombres de etiquetas del XML de Word.
from docx.shared import Mm, Pt                 # Medidas: milímetros (Mm) y puntos tipográficos (Pt).

import graficos                                # Los seis gráficos del informe (graficos.py).
import modelo                                  # Todo el cálculo financiero (modelo.py).

VERSION_SERVIDOR = "3.0"                # Versión del servidor; aparece en / y en /health.
MAX_INFORMES = 500                      # informes que se conservan en memoria
LIMA = timezone(timedelta(hours=-5))    # hora de Lima para el nombre del archivo (UTC−5)

# Tres «cajones» en la memoria del servidor. Se vacían cada vez que Railway reinicia el servidor.
INFORMES = {}          # id del informe -> (nombre_archivo, texto, parámetros del análisis)
ANALISIS = {}          # último análisis por empresa: (monto, perfil, pérdida máxima)
IMAGENES = {}          # gráficos ya dibujados por informe: id -> lista de gráficos

# Creación del servidor. El título, la versión y la descripción aparecen en la página /docs.
app = FastAPI(title="Asesor de Portafolios Inteligentes",
              version=VERSION_SERVIDOR,
              description="Modelo del TAF del curso 012755 · CENTRUM PUCP")


# Estas dos «clases» describen la forma exacta de los datos que puede recibir el servidor.
# Si el agente envía algo distinto (por ejemplo, un monto negativo o sin perfil), FastAPI lo
# rechaza automáticamente con un mensaje 422 que explica qué campo está mal.
#   Field(...)      : los tres puntos indican que el campo es OBLIGATORIO.
#   Field("")/None  : valor que toma el campo si no se envía (es opcional).
#   gt=0            : «greater than» 0, debe ser mayor que cero.
#   description     : explicación del campo en la página /docs.
#   examples        : valor de ejemplo que aparece en /docs.

class Solicitud(BaseModel):
    # Datos que llegan a /analizar.
    monto: float = Field(..., gt=0, description="Monto a invertir en dólares", examples=[100000])
    perfil: str = Field(..., description="conservador, moderado o agresivo", examples=["conservador"])
    empresa: Optional[str] = Field("", description="Nombre de la empresa", examples=["Walker y Wilson"])
    perdida_maxima_usd: Optional[float] = Field(None, description="Pérdida máxima tolerada en un día, en dólares",
                                                examples=[2000])


class Informe(BaseModel):
    # Datos que llegan a /informe: el texto que redactó el agente.
    empresa: Optional[str] = Field("", description="Nombre de la empresa", examples=["Walker y Wilson Company"])
    contenido: str = Field(..., min_length=20, description="Texto completo del informe en formato Markdown")
                                                # min_length=20: al menos 20 caracteres, para no guardar informes vacíos.


def _nombre_archivo(empresa: str) -> str:
    """Informe_Walker_y_Wilson_Company_2026-11-08_1130.md (sin tildes ni símbolos)."""
    base = unicodedata.normalize("NFKD", empresa or "empresa").encode("ascii", "ignore").decode()
                                    # normalize("NFKD") separa cada letra de su tilde ("ñ" → "n" + "~");
                                    # encode("ascii", "ignore") descarta las tildes; decode() vuelve a texto.
                                    # Si no llegó empresa, se usa la palabra "empresa".
    base = re.sub(r"[^A-Za-z0-9]+", "_", base).strip("_")[:60] or "empresa"
                                    # Todo lo que no sea letra o número se cambia por guion bajo
                                    # ("Walker y Wilson Company" → "Walker_y_Wilson_Company"), se quitan los guiones
                                    # de los extremos y se corta a 60 caracteres.
    return f"Informe_{base}_{datetime.now(LIMA):%Y-%m-%d_%H%M}.md"
                                    # Se agregan la fecha y la hora de Lima: año-mes-día_horaminuto.


def _direccion_publica(request: Request) -> str:
    """Dirección https del servidor tal como la ve el agente (Railway usa un proxy)."""
    # Railway recibe las peticiones en un intermediario (proxy) y se las pasa al servidor. Ese intermediario
    # anota la dirección original en las cabeceras «x-forwarded-host» y «x-forwarded-proto».
    host = request.headers.get("x-forwarded-host") or request.headers.get("host") or request.url.netloc
                                    # Nombre del servidor: se prueba primero la cabecera del proxy, luego la normal y
                                    # por último la dirección de la propia petición.
    esquema = request.headers.get("x-forwarded-proto", "https")
                                    # Protocolo: https, salvo que el proxy indique otro.
    if host.startswith(("localhost", "127.0.0.1")):
        esquema = request.url.scheme  # En una prueba local (tu propia computadora) se usa el protocolo real, normalmente http.
    return f"{esquema}://{host}"    # Ejemplo: "https://asesor-portafolios.up.railway.app".



def _clave_empresa(empresa) -> str:
    # Convierte el nombre de la empresa en una clave simple para reconocerla aunque se escriba distinto.
    # Ejemplo: "Walker y Wilson Company" y "WALKER Y WILSON COMPANY" → "walkerywilsoncompany".
    t = unicodedata.normalize("NFKD", empresa or "").encode("ascii", "ignore").decode().lower()
                                    # Sin tildes y en minúsculas.
    return re.sub(r"[^a-z0-9]", "", t)
                                    # Se borra todo lo que no sea letra o número (espacios, puntos, &, ...).


def _parametros_del_texto(texto: str):
    """Si no hubo análisis previo, lee el monto, el perfil y la pérdida máxima del propio informe."""
    # Plan de respaldo: si el servidor se reinició y olvidó el último análisis, busca los datos en el texto
    # del informe para poder dibujar igual los gráficos del Word.
    perfil = re.search(r"\b(conservador|moderado|agresivo)\b", texto, re.I)
                                    # Busca la primera aparición de una de las tres palabras (re.I: sin importar
                                    # mayúsculas). \b marca el inicio y el fin de una palabra completa.
    monto = re.search(r"Monto[^0-9]{0,40}([0-9][0-9 .,\u00a0]*[0-9])", texto)
                                    # Busca la palabra «Monto», hasta 40 caracteres que no sean cifras y luego una cifra
                                    # que puede tener espacios, puntos o comas. Ejemplo: "Monto: USD 100 000" → "100 000".
                                    # \u00a0 es el espacio duro que a veces usa Word entre las cifras.
    if not (perfil and monto):      # Si falta alguno de los dos datos, no se puede continuar.
        return None
    cifra = float(re.sub(r"[^0-9]", "", monto.group(1).split(",")[0]))
                                    # Se descarta lo que viene después de una coma (los decimales), se dejan solo las cifras
                                    # y se convierte en número: "100 000" → 100000.0.
    perdida = re.search(r"m[aá]xim[oa] tolerad[oa][^0-9]{0,30}([0-9][0-9 .\u00a0]*[0-9])", texto, re.I)
                                    # Busca «máxima tolerada» (con o sin tilde, en masculino o femenino) y la cifra que sigue.
                                    # Ejemplo: "Pérdida máxima tolerada: USD 2 000" → "2 000".
    tope = float(re.sub(r"[^0-9]", "", perdida.group(1))) if perdida else None
                                    # Se convierte en número; si el informe no la menciona, queda en None.
    return (cifra, perfil.group(1).lower(), tope) if cifra > 0 else None
                                    # Devuelve (monto, perfil, pérdida máxima), el mismo formato de ANALISIS.

# ----------------------------------------------------------------------------------
# Conversión del informe en Markdown a Word (.docx): títulos (#), listas (- y 1.),
# tablas (| a | b |) y negritas (**texto**). Times New Roman 12, A4, márgenes de 25 mm.
# ----------------------------------------------------------------------------------
# Markdown es un formato de texto simple que usan los agentes: # Título, **negrita**, - viñeta, | tabla |.

def _texto_con_negritas(parrafo, texto: str, tam=None):
    # Escribe el texto en un párrafo de Word respetando las negritas marcadas con **...**.
    # re.split con un grupo entre paréntesis separa el texto y CONSERVA lo que estaba entre asteriscos.
    # Ejemplo: "El **VaR** es bajo" → ["El ", "VaR", " es bajo"]. Las posiciones impares (1, 3, ...)
    # son las que estaban entre asteriscos.
    for i, trozo in enumerate(re.split(r"\*\*(.+?)\*\*", texto)):
        if not trozo:                              # Se saltan los trozos vacíos.
            continue
        r = parrafo.add_run(trozo.replace("`", ""))  # Cada trozo es un «run» (tramo de texto con el mismo formato).
                                                     # Se quitan las comillas invertidas que Markdown usa para código.
        r.bold = (i % 2 == 1)                      # Negrita solo en las posiciones impares.
        if tam:
            r.font.size = Pt(tam)                  # Tamaño de letra distinto si se indica (en las tablas, 10,5 pt).


def _celdas(linea: str):
    # Separa una fila de tabla Markdown en sus celdas. Ejemplo: "| KO | 73,6 % |" → ["KO", "73,6 %"].
    return [c.strip() for c in linea.strip().strip("|").split("|")]


def _sombrear(celda, color="E7E7E7"):
    # Pinta el fondo de una celda de Word con gris claro (E7E7E7). python-docx no trae esta función,
    # por eso se agrega directamente la etiqueta XML <w:shd> (sombreado) a las propiedades de la celda.
    tcPr = celda._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), color)
    tcPr.append(sh)


# Dónde se insertan los gráficos: cuando un título del informe contiene la palabra de la izquierda,
# los gráficos de la derecha se colocan al final de esa sección.
# Ejemplo: en la sección «Reparto del portafolio» va el gráfico del reparto; en la del backtest, la curva y las caídas.
SECCIONES_GRAFICOS = [("reparto", ["reparto"]), ("riesgo", ["riesgo"]),
                      ("fuera de muestra", ["base100", "velas"]), ("backtest", ["backtest", "caidas"])]


def _insertar_graficos(doc, pendientes, imagenes, contador):
    # Inserta en el Word los gráficos que están en espera (pendientes), cada uno con su pie de figura numerado.
    #   imagenes : diccionario clave -> (título, png) con los gráficos aún no insertados.
    #   contador : lista de un elemento con el número de la última figura (se usa una lista para que el número
    #              se conserve entre una llamada y otra).
    for clave in pendientes:
        if clave not in imagenes:                  # Si ese gráfico ya se insertó, se salta.
            continue
        titulo, png = imagenes.pop(clave)          # pop lo saca del diccionario para no insertarlo dos veces.
        contador[0] += 1                           # Número de figura siguiente: 1, 2, 3, ...
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER    # Imagen centrada.
        p.paragraph_format.keep_with_next = True   # La imagen no se separa de su pie de figura en otra página.
        p.paragraph_format.space_before = Pt(6)
        p.add_run().add_picture(io.BytesIO(png), width=Mm(160))
                                                   # Imagen de 160 mm de ancho: el ancho útil de una A4 con márgenes de 25 mm.
        pie = doc.add_paragraph()
        pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = pie.add_run(f"Figura {contador[0]}. {titulo} Fuente: modelo {modelo.VERSION}.")
                                                   # Ejemplo: "Figura 1. Reparto del dinero ... Fuente: modelo 3.0."
        r.italic = True                            # Pie de figura en cursiva...
        r.font.size = Pt(10)                       # ...y en 10 pt.
    pendientes.clear()                             # Se vacía la lista de espera.


def _a_word(texto: str, graficos_informe=None) -> bytes:
    # Convierte el informe en Markdown en un documento de Word y devuelve sus bytes.
    # graficos_informe es la lista de (clave, título, png) que entrega graficos.graficos().
    imagenes = {k: (t, png) for k, t, png in (graficos_informe or [])}
                                                   # Diccionario clave -> (título, png). Si no hay gráficos, queda vacío.
    pendientes, contador = [], [0]                 # Lista de espera de gráficos y contador de figuras en 0.
    doc = Document()                               # Documento de Word en blanco.
    sec = doc.sections[0]                          # Configuración de página:
    sec.page_height, sec.page_width = Mm(297), Mm(210)   # tamaño A4 (210 × 297 mm)...
    sec.top_margin = sec.bottom_margin = sec.left_margin = sec.right_margin = Mm(25)   # ...con márgenes de 25 mm.
    normal = doc.styles["Normal"]                  # Estilo Normal, el que usan todos los párrafos:
    normal.font.name = "Times New Roman"; normal.font.size = Pt(12)                    # Times New Roman 12...
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")                # ...también para otros alfabetos.
    normal.paragraph_format.space_after = Pt(6)    # 6 pt de espacio después de cada párrafo.
    normal.paragraph_format.widow_control = True   # Control de líneas viudas y huérfanas: Word no deja una línea
                                                   # suelta al pie o al inicio de una página.

    lineas = texto.splitlines()                    # El informe se divide en líneas.
    i = 0                                          # Número de la línea que se está leyendo.
    while i < len(lineas):                         # Se recorre el informe línea por línea.
        linea = lineas[i].rstrip()                 # Se quitan los espacios del final.
        limpia = linea.strip()                     # Y también los del inicio.
        if not limpia or re.fullmatch(r"[-*_]{3,}", limpia):
            i += 1                                 # Las líneas vacías y las líneas separadoras (---, ***, ___) se saltan.
            continue
        # Tabla: filas seguidas que empiezan con |
        if limpia.startswith("|"):
            filas = []
            while i < len(lineas) and lineas[i].strip().startswith("|"):
                fila = _celdas(lineas[i])          # Se separan las celdas de la fila.
                if not all(re.fullmatch(r":?-{2,}:?", c or "--") for c in fila):
                    filas.append(fila)             # Se guarda la fila, salvo la de guiones (|---|---|), que en Markdown
                                                   # solo separa la cabecera del cuerpo.
                i += 1
            ncol = max(len(f) for f in filas)      # Número de columnas: el de la fila más larga.
            tabla = doc.add_table(rows=0, cols=ncol)
            tabla.style = "Table Grid"             # Tabla con todos los bordes.
            for n, fila in enumerate(filas):
                celdas = tabla.add_row().cells     # Se agrega una fila en Word.
                for j in range(ncol):
                    p = celdas[j].paragraphs[0]
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.keep_with_next = n < len(filas) - 1
                                                   # Cada fila se mantiene junto a la siguiente: la tabla no se parte
                                                   # entre dos páginas.
                    _texto_con_negritas(p, fila[j] if j < len(fila) else "", tam=10.5)
                                                   # Texto de la celda en 10,5 pt (vacío si la fila tiene menos celdas).
                    if n == 0:                     # La primera fila es la cabecera:
                        for r in p.runs:
                            r.bold = True          # en negrita...
                        _sombrear(celdas[j])       # ...y con fondo gris claro.
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
                                                   # Párrafo vacío pequeño después de la tabla.
            continue
        m = re.match(r"^(#{1,6})\s*(.+)$", limpia)
                                                   # ¿La línea es un título? De 1 a 6 signos # al inicio y luego el texto.
                                                   # Ejemplo: "## Riesgo del portafolio" → nivel 2.
        if m:                                           # títulos
            _insertar_graficos(doc, pendientes, imagenes, contador)
                                                   # Antes de abrir una sección nueva, se insertan los gráficos que
                                                   # quedaron pendientes de la sección anterior.
            encabezado = m.group(2).lower()        # Texto del título en minúsculas.
            for palabra, claves in SECCIONES_GRAFICOS:
                if palabra in encabezado:          # Si el título contiene una palabra clave (por ejemplo, «riesgo»)...
                    pendientes.extend(c for c in claves if c in imagenes)
                                                   # ...sus gráficos quedan en espera hasta que termine la sección.
                    break
            nivel = len(m.group(1))                # Nivel del título = número de signos #.
            p = doc.add_paragraph()
            p.paragraph_format.keep_with_next = True   # El título nunca queda solo al pie de una página.
            p.paragraph_format.space_before = Pt(14 if nivel <= 2 else 8)
            if nivel == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER  # El título principal va centrado.
            r = p.add_run(m.group(2).replace("**", "").strip())
            r.bold = True
            r.font.size = Pt({1: 16, 2: 14}.get(nivel, 12))
                                                   # Tamaño según el nivel: 16 pt, 14 pt y 12 pt para los demás.
            i += 1
            continue
        m = re.match(r"^[-*•]\s+(.+)$", limpia)    # ¿Es una viñeta? Empieza con -, * o • seguido de espacio.
        if m:                                           # viñetas
            p = doc.add_paragraph(style="List Bullet")  # Estilo de lista con viñetas de Word.
            _texto_con_negritas(p, m.group(1))
            i += 1
            continue
        m = re.match(r"^(\d+)[.)]\s+(.+)$", limpia)
                                                   # ¿Es una lista numerada? Ejemplo: "1. Primer punto" o "2) Segundo".
        if m:                                           # listas numeradas
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Mm(6)          # Sangría de 6 mm...
            p.paragraph_format.first_line_indent = Mm(-6)   # ...con sangría francesa: el número sobresale a la izquierda.
            _texto_con_negritas(p, f"{m.group(1)}. {m.group(2)}")
            i += 1
            continue
        p = doc.add_paragraph()                         # párrafo normal
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY        # Texto justificado.
        _texto_con_negritas(p, limpia)
        i += 1

    _insertar_graficos(doc, pendientes, imagenes, contador)
                                                   # Al terminar el texto, se insertan los gráficos de la última sección.
    if imagenes:                                        # gráficos cuya sección no apareció
        # Si el informe no tenía un título con alguna de las palabras clave, sus gráficos se reúnen al final
        # en una sección propia, para que ningún gráfico se quede fuera del Word.
        p = doc.add_paragraph()
        p.paragraph_format.keep_with_next = True
        p.paragraph_format.space_before = Pt(14)
        r = p.add_run("Gráficos del análisis")
        r.bold = True
        r.font.size = Pt(14)
        _insertar_graficos(doc, [k for k, _, _ in (graficos_informe or []) if k in imagenes], imagenes, contador)
                                                   # Se respeta el orden original de los gráficos.
    salida = io.BytesIO()                          # «Archivo» en memoria...
    doc.save(salida)                               # ...donde se guarda el Word...
    return salida.getvalue()                       # ...y se devuelven sus bytes.


@app.get("/")
def inicio():
    # Ruta raíz: presentación del servicio. Al abrir la dirección del servidor en el navegador se ve este JSON.
    # «@app.get("/")» es un decorador: le indica a FastAPI que esta función responde a GET en la dirección /.
    return {"servicio": "Asesor de Portafolios Inteligentes", "version": VERSION_SERVIDOR,
            "version_modelo": modelo.VERSION, "perfiles": list(modelo.PERFILES),
            "rutas": ["/health", "/analizar", "/informe", "/descargar/{id}", "/descargar/{id}/word", "/docs"]}


@app.get("/health")
def salud():
    # Ruta de salud: Railway y el agente la consultan para comprobar que el servidor está en marcha.
    # "sesiones" es el número de días de datos cargados: confirma que precios.csv se leyó bien.
    return {"estado": "ok", "version": VERSION_SERVIDOR, "version_modelo": modelo.VERSION,
            "sesiones": len(modelo.LOG)}


@app.post("/analizar")
def analizar(s: Solicitud):
    # Ruta principal. El agente envía los datos; FastAPI los valida contra la clase Solicitud y los entrega en s.
    try:
        resultado = modelo.analizar(s.monto, s.perfil, s.empresa or "", s.perdida_maxima_usd)
                                                         # Se llama al modelo con los cuatro datos recibidos.
        parametros = (s.monto, s.perfil.strip().lower(), s.perdida_maxima_usd)
        ANALISIS[_clave_empresa(s.empresa)] = parametros # Se recuerdan los datos de esta empresa, para que /informe
                                                         # pueda dibujar después los gráficos del Word.
        ANALISIS[""] = parametros                        # el último, por si el informe no trae empresa
        return resultado                                 # FastAPI convierte el diccionario en JSON y lo envía al agente.
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
                                                         # Si el modelo rechazó un dato (perfil desconocido o monto no
                                                         # positivo), se responde 400 con el mismo mensaje explicativo.


@app.post("/informe")
def guardar_informe(i: Informe, request: Request):
    # Guarda el informe que redactó el agente y devuelve los enlaces para descargarlo.
    if len(INFORMES) >= MAX_INFORMES:                 # descarta el más antiguo
        INFORMES.pop(next(iter(INFORMES)))            # Los diccionarios recuerdan el orden de llegada: el primero es
                                                      # el más antiguo. Así la memoria nunca pasa de 500 informes.
    clave = uuid.uuid4().hex[:12]                     # Identificador al azar de 12 caracteres. Ejemplo: "3f9a1c07b2e4".
                                                      # Nadie puede adivinar el enlace de otro informe.
    nombre = _nombre_archivo(i.empresa or "")         # Nombre del archivo con la empresa, la fecha y la hora.
    INFORMES[clave] = (nombre, i.contenido.strip() + "\n",
                       ANALISIS.get(_clave_empresa(i.empresa)) or _parametros_del_texto(i.contenido)
                       or ANALISIS.get(""))
    # Se guarda (nombre, texto, parámetros). Los parámetros para los gráficos se buscan en este orden:
    #   1. el último análisis de esa misma empresa;
    #   2. si no hay, los que se puedan leer en el propio texto del informe;
    #   3. si tampoco, el último análisis que se hizo, de cualquier empresa.
    enlace = f"{_direccion_publica(request)}/descargar/{clave}"   # Enlace de descarga en .md.
    enlace_word = f"{enlace}/word"                                # Enlace de descarga en Word.
    return {"enlace_descarga": enlace, "enlace_word": enlace_word, "archivo": nombre,
            "archivo_word": nombre[:-3] + ".docx",                # nombre[:-3] quita ".md" para poner ".docx".
            "mensaje": (f"Informe guardado. Descarga el informe en formato .md: {enlace} · "
                        f"Descarga el informe en Word: {enlace_word}")}
                                                      # El agente le muestra este mensaje al usuario.


@app.get("/descargar/{clave}")
def descargar(clave: str):
    # Descarga del informe en .md. {clave} es la parte variable de la dirección: el identificador del informe.
    if clave not in INFORMES:                          # Si el informe no está (por ejemplo, porque el servidor se
        raise HTTPException(status_code=404,           # reinició), se responde 404 con una indicación para el usuario.
                            detail="El informe ya no está disponible. Pide al agente que lo vuelva a guardar.")
    nombre, texto = INFORMES[clave][:2]                # Se toman el nombre y el texto ([:2] = los dos primeros datos).
    return Response(content=texto.encode("utf-8"), media_type="text/markdown; charset=utf-8",
                    headers={"Content-Disposition": f'attachment; filename="{nombre}"'})
                                                       # «attachment» le dice al navegador que lo descargue como archivo
                                                       # con ese nombre, en vez de mostrarlo en pantalla.


@app.get("/descargar/{clave}/word")
def descargar_word(clave: str):
    # Descarga del mismo informe en Word, con los seis gráficos.
    if clave not in INFORMES:
        raise HTTPException(status_code=404,
                            detail="El informe ya no está disponible. Pide al agente que lo vuelva a guardar.")
    nombre, texto, parametros = INFORMES[clave]
    if clave not in IMAGENES:                          # Los gráficos se dibujan solo la primera vez que se pide el Word;
        try:                                           # las siguientes descargas reutilizan los ya dibujados.
            IMAGENES[clave] = graficos.graficos(*parametros) if parametros else []
                                                       # *parametros reparte la tupla (monto, perfil, pérdida) como los
                                                       # tres datos de la función. Sin parámetros, no hay gráficos.
        except Exception:                               # sin gráficos, pero el Word igual se entrega
            IMAGENES[clave] = []
        if len(IMAGENES) > 50:                         # Se conservan como máximo 50 juegos de gráficos en memoria:
            IMAGENES.pop(next(iter(IMAGENES)))         # se descarta el más antiguo.
    return Response(content=_a_word(texto, IMAGENES[clave]),
                    media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    headers={"Content-Disposition": f'attachment; filename="{nombre[:-3]}.docx"'})
                                                       # media_type es el tipo oficial de un archivo .docx; el navegador lo
                                                       # reconoce y ofrece abrirlo con Word.
