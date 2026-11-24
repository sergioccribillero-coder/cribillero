# -*- coding: utf-8 -*-
"""
Asesor de Portafolios Inteligentes - gráficos del informe en Word  ·  versión 3.0
Curso 012755 Construcción de Portafolios Inteligentes · CENTRUM PUCP

Dibuja, con los mismos datos y funciones de modelo.py, los seis gráficos que el
servidor inserta en el informe de Word. Estilo TradingView sobre fondo blanco:
verde para las subidas, rojo para las bajadas, eje de precios a la derecha.

  1. Reparto del dinero (barras horizontales)
  2. Velas mensuales del S&P 500 con medias móviles de 50 y 200 sesiones
  3. Comparación de los activos en base 100 (escala logarítmica)
  4. Distribución de las ganancias y pérdidas diarias con el VaR al 95 %
  5. Curva del backtest: método elegido, pesos iguales y S&P 500
  6. Caídas desde el máximo del portafolio y del S&P 500

Historial:
  1.0  Primera versión de los gráficos.
  1.1  Cada línea lleva un comentario explicativo. Los gráficos son idénticos a la 1.0.
  3.0  Numeración unificada con el material del curso. Los gráficos son idénticos a la 1.1.
"""
import io                              # io: permite guardar la imagen en la memoria, sin crear archivos en el disco.

import matplotlib                      # matplotlib: la biblioteca de gráficos de Python.
matplotlib.use("Agg")                  # «Agg» dibuja sin pantalla. Es indispensable en Railway, que es un servidor
                                       # sin monitor: los gráficos se generan directamente como imágenes PNG.
import matplotlib.pyplot as plt        # pyplot (alias plt): las funciones para crear figuras y ejes.
import matplotlib.dates as mdates      # mdates: formato de las fechas del eje horizontal (2015, 2016, ...).
from matplotlib.ticker import FuncFormatter
                                       # FuncFormatter: permite escribir las cifras de los ejes con una función propia.
import numpy as np                     # NumPy: cálculo numérico.
import pandas as pd                    # pandas: tablas con fechas.

import modelo as m                     # Se importa el modelo con el alias m: los gráficos usan sus mismos datos y
                                       # funciones, de modo que las cifras del Word coinciden con las del JSON.

VERSION = "3.0"                        # Versión de este archivo.

# Colores al estilo TradingView, en código hexadecimal (#RRGGBB: rojo, verde y azul de 00 a ff).
VERDE, ROJO, AZUL, NARANJA, MORADO, GRIS = "#26a69a", "#ef5350", "#2962ff", "#ff9800", "#7e57c2", "#787b86"

# Un color fijo por activo, para que cada uno se reconozca igual en todos los gráficos.
COLORES = {"NFLX": "#e53935", "AMD": "#00897b", "HD": "#6d4c41", "KO": "#8e24aa",
           "BTC-USD": "#f9a825", "^GSPC": AZUL}

# Nombre comercial de cada ticker, para las leyendas y etiquetas.
NOMBRES = {"NFLX": "Netflix", "AMD": "AMD", "HD": "Home Depot", "KO": "Coca-Cola",
           "BTC-USD": "Bitcoin", "^GSPC": "S&P 500"}

# Estilo general de todos los gráficos (se fija una sola vez):
plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 9,        # Tipografía y tamaño de letra.
    "axes.edgecolor": "#d1d4dc",                         # Color del borde de los ejes (gris claro).
    "axes.labelcolor": "#131722",                        # Color de los títulos de los ejes (casi negro).
    "xtick.color": "#434651", "ytick.color": "#434651",  # Color de las cifras de los ejes.
    "axes.grid": True, "grid.color": "#eef0f3", "grid.linewidth": 0.8,   # Cuadrícula tenue de fondo.
    "axes.spines.top": False, "axes.spines.left": False, # Sin línea superior ni izquierda (el eje va a la derecha).
    "figure.dpi": 100,                                   # Resolución de trabajo de la figura.
})


# Funciones auxiliares. El guion bajo inicial indica que son de uso interno de este archivo.

def _miles(x, _=None):
    # Formato de miles con espacio para los ejes. Ejemplo: 307519 → "307 519".
    # El segundo parámetro (_) lo exige FuncFormatter, pero no se usa.
    return f"{x:,.0f}".replace(",", " ")


def _pct(x, _=None):
    # Formato de porcentaje sin decimales para los ejes. Ejemplo: −25.3 → "-25 %".
    return f"{x:.0f} %"


def _figura(alto=3.6):
    # Crea una figura de 6,6 pulgadas de ancho (cabe en una página A4 con márgenes) y el alto indicado.
    fig, ax = plt.subplots(figsize=(6.6, alto))      # fig es la hoja completa; ax, el área donde se dibuja.
    ax.yaxis.tick_right()                            # Cifras del eje vertical a la derecha, como en TradingView.
    ax.yaxis.set_label_position("right")             # El título de ese eje también va a la derecha.
    return fig, ax


def _png(fig):
    # Convierte la figura en una imagen PNG guardada en la memoria y devuelve sus bytes.
    buf = io.BytesIO()                               # «Archivo» temporal en la memoria.
    fig.tight_layout()                               # Ajusta márgenes para que no se corten títulos ni etiquetas.
    fig.savefig(buf, format="png", dpi=200, facecolor="white")
                                                     # Guarda en PNG a 200 puntos por pulgada (buena calidad de impresión)
                                                     # con fondo blanco.
    plt.close(fig)                                   # Cierra la figura para liberar memoria en el servidor.
    return buf.getvalue()                            # Bytes de la imagen, listos para insertarlos en el Word.


def _titulo(ax, texto, sub=""):
    # Escribe el título en negrita alineado a la izquierda y, si se indica, un subtítulo gris encima del gráfico.
    ax.set_title(texto, loc="left", fontsize=10.5, fontweight="bold", color="#131722", pad=14)
    if sub:
        ax.text(0, 1.015, sub, transform=ax.transAxes, fontsize=8, color=GRIS, va="bottom")
        # transform=ax.transAxes: las coordenadas (0, 1.015) se miden respecto del área del gráfico,
        # donde 0 es el borde izquierdo y 1 el borde superior. Así el subtítulo queda justo encima.


def graficos(monto, perfil, perdida_maxima_usd=None):
    """Devuelve una lista de (clave, título, png) con los seis gráficos."""
    # Recibe los mismos datos que analizar() en modelo.py, excepto el nombre de la empresa.
    perfil = perfil.strip().lower()               # Limpia el texto del perfil (espacios y mayúsculas).
    metodo = m.PERFILES[perfil]["metodo"]         # Método del perfil, tomado del modelo.
    activos = m.PERFILES[perfil]["activos"]       # Activos del perfil.
    w = m.METODOS[metodo](m.LOG[activos])         # Pesos del portafolio: el mismo cálculo del reparto en modelo.py.
    salida = []                                   # Lista donde se guardará cada gráfico como (clave, título, imagen).

    # 1. Reparto del dinero -------------------------------------------------
    orden = [i for i in np.argsort(w) if w[i] >= 0.0005]
                                       # Activos de menor a mayor peso, sin los de peso casi nulo (menos de 0,05 %).
                                       # Se ordena de menor a mayor porque las barras horizontales se dibujan de abajo
                                       # hacia arriba: así el activo con más peso queda arriba.
    fig, ax = _figura(0.55 * len(orden) + 1.4)
                                       # El alto de la figura crece con el número de barras (0,55 pulgadas por barra).
    etiquetas = [f"{NOMBRES[activos[i]]} ({activos[i]})" for i in orden]
                                       # Texto de cada barra. Ejemplo: "Coca-Cola (KO)".
    dolares = [w[i] * monto for i in orden]
                                       # Largo de cada barra: el dinero asignado. Ejemplo: 0,7356 × 100 000 = 73 558.
    barras = ax.barh(etiquetas, dolares, color=[COLORES[activos[i]] for i in orden], height=0.6)
                                       # barh dibuja barras horizontales, cada una con el color de su activo.
    ax.yaxis.tick_left()               # En este gráfico los nombres van a la izquierda, para leerlos mejor.
    ax.xaxis.set_major_formatter(FuncFormatter(_miles))
                                       # Cifras del eje horizontal con espacio de miles.
    ax.set_xlabel("Dólares")
    for b, i in zip(barras, orden):    # Junto a cada barra se escribe su porcentaje y su monto.
        ax.text(b.get_width() + monto * 0.01, b.get_y() + b.get_height() / 2,
                f"{m.pct(w[i])} · {m.usd(w[i] * monto)}", va="center", fontsize=8.5)
                                       # Posición: al final de la barra más un 1 % del monto de separación,
                                       # centrado en la altura de la barra. Ejemplo: "73,6 % · USD 73 558".
    ax.set_xlim(0, max(dolares) * 1.35)
                                       # El eje llega a 1,35 veces la barra más larga, para que quepa el texto.
    ax.grid(axis="y", visible=False)   # Sin líneas horizontales de cuadrícula entre barras.
    _titulo(ax, f"Reparto de {m.usd(monto)} · perfil {perfil}", f"Método: {metodo} · modelo {m.VERSION}")
    salida.append(("reparto", "Reparto del dinero entre los activos del portafolio recomendado.", _png(fig)))
                                       # Se guarda el gráfico con su clave («reparto»), su título para el pie de
                                       # figura del Word y la imagen PNG.

    # 2. Velas mensuales del S&P 500 -----------------------------------------
    sp = m.PRECIOS[m.INDICE]                       # Serie de precios diarios del S&P 500.
    velas = sp.resample("ME").agg(["first", "max", "min", "last"])
                                                   # resample("ME") agrupa los días por mes (ME = fin de mes) y agg calcula
                                                   # cuatro datos de cada mes: primer cierre, el mayor, el menor y el último.
                                                   # Como la base solo tiene cierres diarios, cada vela se arma con cierres.
    velas.columns = ["apertura", "maximo", "minimo", "cierre"]
                                                   # Se renombran las columnas con los nombres de una vela japonesa.
    sma50, sma200 = sp.rolling(50).mean().resample("ME").last(), sp.rolling(200).mean().resample("ME").last()
                                                   # Medias móviles simples (SMA, Simple Moving Average) de 50 y 200 sesiones:
                                                   # el promedio de los últimos 50 o 200 cierres, día a día. Se toma el valor
                                                   # del último día de cada mes para dibujarlas junto a las velas mensuales.
    fig, ax = _figura(3.8)
    ancho = 20                                     # Ancho de cada vela, en días.
    for fecha, f in velas.iterrows():              # Se dibuja una vela por mes.
        color = VERDE if f.cierre >= f.apertura else ROJO
                                                   # Verde si el mes cerró por encima de donde empezó; rojo si cerró por debajo.
        ax.vlines(fecha, f.minimo, f.maximo, color=color, linewidth=0.8)
                                                   # Mecha: línea vertical del mínimo al máximo del mes.
        ax.bar(fecha, abs(f.cierre - f.apertura) or sp.iloc[-1] * 0.002, bottom=min(f.apertura, f.cierre),
               width=ancho, color=color, edgecolor=color, linewidth=0.5)
                                                   # Cuerpo: rectángulo entre la apertura y el cierre. Si ambos coinciden
                                                   # (diferencia 0), «or» le da un alto mínimo de 0,2 % del último precio
                                                   # para que la vela se vea.
    ax.plot(sma50.index, sma50.values, color=NARANJA, linewidth=1.2, label="Media móvil 50 sesiones")
    ax.plot(sma200.index, sma200.values, color=MORADO, linewidth=1.2, label="Media móvil 200 sesiones")
    ax.yaxis.set_major_formatter(FuncFormatter(_miles))       # Cifras del eje con espacio de miles.
    ax.xaxis.set_major_locator(mdates.YearLocator())          # Una marca por año en el eje horizontal...
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))  # ...escrita solo con el año: 2015, 2016, ...
    ax.legend(loc="upper left", frameon=False, fontsize=8)    # Leyenda arriba a la izquierda, sin recuadro.
    _titulo(ax, "S&P 500 · velas mensuales", "Cada vela: primer, mayor, menor y último cierre del mes · 2015-2024")
    salida.append(("velas", "Velas mensuales del S&P 500 con sus medias móviles de 50 y 200 sesiones.", _png(fig)))

    # 3. Comparación en base 100 ---------------------------------------------
    cols = m.ACCIONES + m.CRIPTO + [m.INDICE]               # Los seis activos: cuatro acciones, bitcoin y S&P 500.
    base = m.PRECIOS[cols] / m.PRECIOS[cols].iloc[0] * 100  # Base 100: cada precio se divide entre su primer precio y se
                                                            # multiplica por 100. Todos arrancan en 100 y se comparan entre
                                                            # sí. Ejemplo: si KO termina en 180, 100 dólares se volvieron 180.
    fig, ax = _figura(3.8)
    for c in cols:                                          # Una línea por activo.
        ax.plot(base.index, base[c], color=COLORES[c], linewidth=1.6 if c == m.INDICE else 1.0,
                label=f"{NOMBRES[c]} ({m.num(base[c].iloc[-1] / 100, 1)}×)")
                                                            # El S&P 500 va con línea más gruesa. En la leyenda se indica
                                                            # cuántas veces se multiplicó el dinero. Ejemplo: "Coca-Cola (1,8×)".
    ax.set_yscale("log")                                    # Escala logarítmica: la misma distancia vertical significa el mismo
                                                            # porcentaje de subida. Sin ella, el bitcoin aplastaría a los demás.
    ax.yaxis.set_major_formatter(FuncFormatter(_miles))
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.legend(loc="upper left", frameon=False, fontsize=7.5, ncol=2)   # Leyenda en dos columnas.
    _titulo(ax, "Comparación de los activos · base 100", "Valor de 100 dólares invertidos el 02/01/2015 · escala logarítmica")
    salida.append(("base100", "Evolución de cada activo y del S&P 500 desde una inversión de 100 dólares en 2015.", _png(fig)))

    # 4. Distribución de las ganancias y pérdidas diarias --------------------
    r = m.serie_portafolio(w, activos)             # Rendimiento logarítmico diario del portafolio, 2015-2024.
    usd_dia = (np.exp(r) - 1) * monto              # Se convierte en dólares ganados o perdidos cada día.
                                                   # Ejemplo: r = −0,01 → (e^−0,01 − 1) × 100 000 = −995 dólares.
    var95, _ = m.var_historico(r.values, 0.95)     # VaR (Valor en Riesgo) al 95 % a un día, el mismo del modelo.
    fig, ax = _figura(3.4)
    n, bordes, parches = ax.hist(usd_dia, bins=80, color=VERDE, alpha=0.85)
                                                   # Histograma: divide el rango de resultados diarios en 80 tramos y cuenta
                                                   # cuántos días cayeron en cada uno. Devuelve los conteos (n), los bordes
                                                   # de los tramos y las barras dibujadas (parches).
    for p, izq in zip(parches, bordes[:-1]):       # Se revisa cada barra junto con su borde izquierdo.
        if izq < -var95 * monto:                   # Si la barra está a la izquierda del VaR (días peores que el VaR)...
            p.set_facecolor(ROJO)                  # ...se pinta de rojo: es el 5 % de los días más desfavorables.
    ax.axvline(-var95 * monto, color=ROJO, linestyle="--", linewidth=1.3,
               label=f"VaR 95 % a un día: {m.usd(var95 * monto)}")
                                                   # Línea vertical discontinua en el VaR.
    if perdida_maxima_usd:                         # Si el cliente fijó una pérdida máxima tolerada...
        ax.axvline(-perdida_maxima_usd, color="#131722", linestyle=":", linewidth=1.3,
                   label=f"Pérdida máxima tolerada: {m.usd(perdida_maxima_usd)}")
                                                   # ...se dibuja con una línea punteada negra. Si queda a la izquierda de
                                                   # la línea roja, el portafolio respeta el límite.
    ax.xaxis.set_major_formatter(FuncFormatter(_miles))
    ax.set_xlabel("Ganancia o pérdida de un día (dólares)")
    ax.set_ylabel("Número de días")
    ax.legend(loc="upper left", frameon=False, fontsize=8)
    _titulo(ax, "Ganancias y pérdidas diarias del portafolio", "2015-2024 · en rojo, los días peores que el VaR al 95 %")
    salida.append(("riesgo", "Distribución de las ganancias y pérdidas diarias del portafolio, con el VaR al 95 %.", _png(fig)))

    # 5. Curva del backtest ---------------------------------------------------
    v, _ = m.backtest(tuple(activos), metodo)                # Curva del método del perfil (memorizada: no se recalcula).
    v_ig, _ = m.backtest(tuple(m.ACCIONES), "Pesos iguales") # Curva de referencia con pesos iguales.
    sp_bt = sp.loc[v.index[0]:]                              # S&P 500 desde el mismo día en que empieza el backtest.
    curvas = {f"{metodo}": v / v.iloc[0] * monto,
              "Pesos iguales": v_ig / v_ig.iloc[0] * monto,
              "S&P 500": sp_bt / sp_bt.iloc[0] * monto}
                                                             # Las tres curvas se pasan a dólares: todas empiezan en el monto
                                                             # invertido. En el perfil moderado el método ya es «Pesos
                                                             # iguales», así que las dos primeras curvas son la misma y el
                                                             # diccionario guarda una sola.
    fig, ax = _figura(3.8)
    colores = {metodo: AZUL, "Pesos iguales": NARANJA, "S&P 500": GRIS}
                                                             # Azul para el método elegido, naranja y gris para las referencias.
    if metodo == "Pesos iguales":                            # En el perfil moderado, la línea de pesos iguales es la del
        colores[metodo] = AZUL                               # método elegido, y por eso va en azul.
    for nombre, serie in curvas.items():
        ax.plot(serie.index, serie.values, color=colores.get(nombre, AZUL), linewidth=1.5 if nombre == metodo else 1.0,
                label=f"{nombre}: {m.usd(serie.iloc[-1])}")
                                                             # Cada curva con su valor final en la leyenda; la del método elegido,
                                                             # más gruesa. Ejemplo: "Mínima varianza: USD 307 519".
    elegida = curvas[metodo]
    ax.fill_between(elegida.index, monto, elegida.values, color=AZUL, alpha=0.08)
                                                             # Sombrea en azul tenue el área entre el monto inicial y la curva
                                                             # elegida: lo que se ganó (o se perdió) en cada momento.
    ax.axhline(monto, color="#b2b5be", linewidth=0.8)        # Línea horizontal en el monto invertido, como referencia.
    ax.yaxis.set_major_formatter(FuncFormatter(_miles))
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.legend(loc="upper left", frameon=False, fontsize=8)
    _titulo(ax, f"Backtest de {m.usd(monto)}", f"{v.index[0]:%d/%m/%Y} al {v.index[-1]:%d/%m/%Y} · rebalanceo mensual y costo de 0,1 %")
    salida.append(("backtest", "Crecimiento del monto invertido con el método elegido, con pesos iguales y con el S&P 500.", _png(fig)))

    # 6. Caídas desde el máximo -----------------------------------------------
    caida_p = (v / v.cummax() - 1) * 100       # Para cada día, cuánto está el portafolio por debajo de su máximo anterior,
                                               # en porcentaje. 0 % significa que ese día marcó un nuevo máximo.
                                               # Ejemplo: máximo previo 1,20 y valor de hoy 0,96 → 0,96 / 1,20 − 1 = −20 %.
    caida_s = (sp_bt / sp_bt.cummax() - 1) * 100
                                               # Lo mismo para el S&P 500.
    fig, ax = _figura(3.2)
    ax.fill_between(caida_p.index, caida_p.values, 0, color=ROJO, alpha=0.35,
                    label=f"{metodo}: peor caída {m.num(caida_p.min(), 1)} %")
                                               # Área roja entre la caída del portafolio y el 0 %; en la leyenda, su peor caída.
    ax.plot(caida_s.index, caida_s.values, color=GRIS, linewidth=1.0,
            label=f"S&P 500: peor caída {m.num(caida_s.min(), 1)} %")
                                               # Línea gris con las caídas del S&P 500, para comparar.
    ax.yaxis.set_major_formatter(FuncFormatter(_pct))
    ax.xaxis.set_major_locator(mdates.YearLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax.legend(loc="lower left", frameon=False, fontsize=8)
    _titulo(ax, "Caídas desde el máximo", "Cuánto está el valor por debajo de su máximo anterior")
    salida.append(("caidas", "Caídas desde el máximo del portafolio del backtest y del S&P 500.", _png(fig)))
    return salida                              # Lista con los seis gráficos, en el orden en que se dibujaron.


# Prueba del archivo: solo se ejecuta al correrlo directamente («python graficos.py»),
# no cuando el servidor lo importa. Dibuja los seis gráficos del perfil conservador,
# los guarda en la carpeta temporal /tmp y muestra cuántos son y cuántos segundos tardó.
if __name__ == "__main__":
    import time
    t = time.time()                                  # Hora de inicio, en segundos.
    g = graficos(100000, "conservador", 2000)
    for k, titulo, png in g:
        open(f"/tmp/g_{k}.png", "wb").write(png)     # "wb": escribe en modo binario, porque un PNG no es texto.
    print(len(g), round(time.time() - t, 1), "s")    # Ejemplo de salida: 6 4.2 s
