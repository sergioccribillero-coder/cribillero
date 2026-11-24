# -*- coding: utf-8 -*-
"""
Asesor de Portafolios Inteligentes - modelo del TAF  ·  versión 3.0
Curso 012755 Construcción de Portafolios Inteligentes · CENTRUM PUCP

Este archivo contiene TODO el cálculo financiero. El servidor (app.py) solo
recibe los datos del agente, llama a analizar() y devuelve el resultado.

Convenciones del TAF:
  - Precios de cierre ajustados de Yahoo Finance, 01/01/2015 - 31/12/2024.
  - Rendimientos logarítmicos diarios; año de 252 sesiones; tasa libre de riesgo 2 %.
  - Prueba fuera de muestra: se estima con 2015-2019 y se evalúa en 2020-2024.
  - Backtest: ventana móvil de 252 sesiones, rebalanceo el primer día hábil de
    cada mes y costo de 0,1 % por cada dólar que se mueve.

Historial:
  1.0  Primera versión del modelo.
  1.1  Cada línea lleva un comentario explicativo. Los cálculos son idénticos a la 1.0.
  3.0  Numeración unificada con el material del curso. Los cálculos son idénticos a la 1.1.
"""
# La línea «# -*- coding: utf-8 -*-» declara que el archivo usa la codificación UTF-8,
# para que las tildes y la letra ñ se lean bien en cualquier computadora.
# El texto entre triples comillas es la «cadena de documentación» (docstring) del archivo:
# Python la guarda como descripción del módulo y no la ejecuta.

# ---------------------------------------------------------------------------
# Importaciones: herramientas que el modelo toma prestadas de otras bibliotecas
# ---------------------------------------------------------------------------
from functools import lru_cache        # lru_cache «memoriza» el resultado de una función: si se vuelve a pedir
                                       # con los mismos datos, entrega la respuesta guardada sin recalcular.
                                       # Se usa en el backtest, que es el cálculo más lento.
from pathlib import Path               # Path arma rutas de archivos que funcionan en Windows, Linux o Mac.

import numpy as np                     # NumPy: cálculo numérico con vectores y matrices (alias np).
import pandas as pd                    # pandas: tablas de datos con fechas (alias pd).
from scipy.optimize import minimize    # minimize: busca los valores que hacen mínima una función (el optimizador).
from scipy.stats import norm           # norm: la distribución normal, que usa el VaR (Valor en Riesgo) paramétrico.

VERSION = "3.0"                        # Versión del modelo. Viaja en cada respuesta del servidor y en los pies de
                                       # figura del Word, para saber con qué versión se calculó cada informe.

# ----------------------------------------------------------------------------
# 1. Parámetros del modelo (cada grupo cambia aquí sus activos)
# ----------------------------------------------------------------------------
ACCIONES = ["NFLX", "AMD", "HD", "KO"]  # Códigos bursátiles (tickers) de las cuatro acciones del portafolio:
                                        # Netflix, Advanced Micro Devices, Home Depot y Coca-Cola.
                                        # Si un grupo del TAF trabaja con otras acciones, solo cambia esta lista
                                        # (y descarga sus precios en la Parte 1).
CRIPTO = ["BTC-USD"]                    # El bitcoin cotizado en dólares. Va en una lista aparte porque el perfil
                                        # moderado no lo puede usar.
INDICE = "^GSPC"                        # Código del índice S&P 500, que reúne a las 500 empresas más grandes de
                                        # Estados Unidos. No se invierte en él: sirve de punto de comparación.
TASA_LIBRE = 0.02                       # Tasa libre de riesgo: 2 % anual. Es lo que ganarías sin arriesgar nada, por
                                        # ejemplo con un bono del Tesoro de Estados Unidos. Se resta en el ratio de Sharpe.
SESIONES = 252                          # Número de días de bolsa de un año, una vez descontados fines de semana y feriados.
                                        # Sirve para ANUALIZAR: un rendimiento diario promedio de 0,05 % equivale a
                                        # 0,05 % × 252 = 12,6 % al año.
VENTANA = 252                           # Cantidad de días de historia que el backtest mira para estimar los pesos, o sea,
                                        # un año de datos. Ejemplo: para decidir los pesos de hoy, el modelo usa solo los
                                        # 252 días anteriores y no toda la serie, así que la estimación se va actualizando
                                        # conforme pasa el tiempo.
                                        # Coincide con SESIONES en el número, pero cumple otra función: SESIONES sirve
                                        # para anualizar y VENTANA define cuánta historia se mira.
COSTO = 0.001                           # Costo de transacción: la comisión que pagas cada vez que compras o vendes un
                                        # activo para armar o reajustar el portafolio. 0.001 equivale a 0,1 % del monto
                                        # negociado, lo que en el mercado se llama 10 puntos básicos.
                                        # Ejemplo: mover USD 10 000 de una acción a otra cuesta USD 10.
CORTE_MUESTRA = "2019-12-31"            # Fecha que divide la historia en dos para la prueba fuera de muestra:
                                        # 2015-2019 para estimar los pesos y 2020-2024 para ver cómo les fue.

# Regla del perfil (la misma del agente de la clase 1)
# Este diccionario asocia cada perfil de riesgo con un método de optimización y con los activos que puede usar:
#
#   Perfil        Método            Activos              Qué hace
#   -----------   ---------------   ------------------   -----------------------------------------------------------
#   Conservador   Mínima varianza   Acciones + bitcoin   Busca la combinación de pesos con la menor varianza posible.
#                                                        La varianza es la desviación estándar al cuadrado, así que es
#                                                        otra forma de medir el riesgo. Puede incluir bitcoin, aunque
#                                                        normalmente le dará un peso muy pequeño.
#   Moderado      Pesos iguales     Solo acciones        Reparte el dinero en partes iguales, 25 % en cada acción, sin
#                                                        optimizar nada. Excluye el bitcoin.
#   Agresivo      Máximo Sharpe     Acciones + bitcoin   Busca los pesos que dan el mayor ratio de Sharpe, es decir, el
#                                                        mejor rendimiento por unidad de riesgo.
#
# Cómo se lee: PERFILES["conservador"]["metodo"] devuelve "Mínima varianza" y
# PERFILES["conservador"]["activos"] devuelve ["NFLX", "AMD", "HD", "KO", "BTC-USD"].
# ACCIONES + CRIPTO une las dos listas en una sola.
PERFILES = {
    "conservador": {"metodo": "Mínima varianza", "activos": ACCIONES + CRIPTO},
    "moderado": {"metodo": "Pesos iguales", "activos": ACCIONES},
    "agresivo": {"metodo": "Máximo Sharpe", "activos": ACCIONES + CRIPTO},
}

# ----------------------------------------------------------------------------
# 2. Datos
# ----------------------------------------------------------------------------
# __file__ es la ruta de este mismo archivo (modelo.py). .parent es su carpeta.
# Así, precios.csv se busca siempre junto a modelo.py, ya sea en Colab o en Railway.
ARCHIVO = Path(__file__).parent / "precios.csv"

# Lectura de los precios:
#   index_col="Date", parse_dates=True : las fechas forman el índice y se reconocen como fechas.
#   .loc["2015-01-01":"2024-12-31"]     : se queda solo con el periodo del TAF (ambos extremos incluidos).
#   .dropna()                           : elimina los días en que falta algún precio (por ejemplo, los fines
#                                         de semana, cuando solo cotiza el bitcoin). Quedan los días de bolsa.
PRECIOS = (pd.read_csv(ARCHIVO, index_col="Date", parse_dates=True)
           .loc["2015-01-01":"2024-12-31"].dropna())

# Rendimientos logarítmicos diarios: ln(precio de hoy / precio de ayer).
#   PRECIOS.shift(1) desplaza la tabla un día hacia abajo, de modo que cada fila queda frente al precio de ayer.
#   El primer día no tiene «ayer» y queda vacío; .dropna() lo elimina.
# Ejemplo: si KO pasa de 100 a 102 dólares, el rendimiento simple es 2 % y el logarítmico ln(1,02) = 1,98 %.
# Se usan porque se pueden SUMAR en el tiempo: la suma de los rendimientos logarítmicos de un año es el
# rendimiento logarítmico del año completo. Son la base de los pesos, del riesgo y de la prueba fuera de muestra.
LOG = np.log(PRECIOS / PRECIOS.shift(1)).dropna()      # rendimientos logarítmicos

# Rendimientos simples diarios: (precio de hoy / precio de ayer) − 1. En el ejemplo anterior, 2 %.
# El backtest los usa porque valorizan exactamente la cartera en dólares, día por día.
SIMPLE = PRECIOS.pct_change().dropna()                  # para valorizar el backtest


# ----------------------------------------------------------------------------
# 3. Métodos de optimización
# ----------------------------------------------------------------------------
# Notación que se usa en todo el archivo:
#   w : vector de pesos del portafolio. Ejemplo: [0.25, 0.25, 0.25, 0.25] = 25 % en cada acción.
#   r : tabla de rendimientos logarítmicos diarios (una columna por activo, una fila por día).

def estadisticas(w, r):
    """Rendimiento, volatilidad y Sharpe anuales de los pesos w."""
    media = r.mean().values * SESIONES       # Rendimiento promedio diario de cada activo × 252 = rendimiento anual
                                             # esperado de cada activo. .values lo convierte en un vector de NumPy.
    cov = r.cov().values * SESIONES          # Matriz de covarianzas anualizada: en la diagonal está la varianza de cada
                                             # activo y fuera de ella, cuánto se mueven juntos cada par de activos.
    rend = float(w @ media)                  # En Python, con NumPy, el símbolo @ multiplica vectores y matrices.
                                             # w @ media = suma de (peso × rendimiento) = rendimiento anual del portafolio.
                                             # Ejemplo: 50 % en un activo que rinde 10 % y 50 % en otro que rinde 20 %
                                             # → 0,5 × 10 % + 0,5 × 20 % = 15 %.
    vol = float(np.sqrt(w @ cov @ w))        # Volatilidad anual del portafolio (desviación estándar) = raíz de wᵀ·Σ·w,
                                             # donde Σ es la matriz de covarianzas. Tiene en cuenta que los activos se
                                             # compensan entre sí: por eso diversificar reduce el riesgo.
    return rend, vol, (rend - TASA_LIBRE) / vol
                                             # El tercer valor es el ratio de Sharpe: (rendimiento − tasa libre) / volatilidad,
                                             # o sea, cuánto rendimiento extra se gana por cada punto de riesgo.
                                             # Ejemplo: (15 % − 2 %) / 20 % = 0,65.


def _optimizar(objetivo, r):
    # Motor común de los métodos que optimizan. Recibe la función «objetivo» que se quiere hacer mínima
    # y devuelve los pesos que lo logran. El guion bajo inicial (_optimizar) indica que es de uso interno.
    n = r.shape[1]                                       # Número de activos (columnas de r). Ejemplo: 5.
    res = minimize(objetivo, np.ones(n) / n, method="SLSQP",
                   bounds=[(0, 1)] * n,
                   constraints={"type": "eq", "fun": lambda w: w.sum() - 1})
    # minimize prueba combinaciones de pesos hasta encontrar la que da el menor valor de «objetivo».
    #   np.ones(n) / n        : punto de partida, pesos iguales (con 5 activos, 20 % cada uno).
    #   method="SLSQP"        : programación cuadrática secuencial por mínimos cuadrados (Sequential Least
    #                           Squares Programming), un método que acepta límites y condiciones.
    #   bounds=[(0, 1)] * n   : cada peso debe estar entre 0 % y 100 %. Al no admitir pesos negativos,
    #                           el modelo no hace ventas en corto (vender lo que no se tiene).
    #   constraints           : condición de igualdad ("eq"): la suma de los pesos menos 1 debe dar 0,
    #                           es decir, se invierte el 100 % del dinero.
    w = np.clip(res.x, 0, None)                          # res.x son los pesos encontrados. clip convierte en 0 cualquier
                                                         # residuo negativo diminuto del cálculo (por ejemplo, −0,0000001).
    return w / w.sum()                                   # Se vuelve a dividir entre la suma para que totalicen exactamente 100 %.


def maximo_sharpe(r):
    # Perfil agresivo. minimize solo sabe buscar mínimos; para obtener el MÁXIMO Sharpe se minimiza
    # el Sharpe con signo negativo. estadisticas(w, r)[2] es el tercer valor que devuelve: el Sharpe.
    # «lambda w: ...» define una función corta sin nombre que recibe los pesos w.
    return _optimizar(lambda w: -estadisticas(w, r)[2], r)


def minima_varianza(r):
    # Perfil conservador. Minimiza la volatilidad, estadisticas(w, r)[1]. Como la volatilidad es la raíz
    # de la varianza, los pesos que minimizan una minimizan también la otra.
    return _optimizar(lambda w: estadisticas(w, r)[1], r)


def pesos_iguales(r):
    # Perfil moderado. No optimiza: reparte por igual. Con 4 acciones: [0.25, 0.25, 0.25, 0.25].
    return np.ones(r.shape[1]) / r.shape[1]


# Diccionario que asocia el nombre de cada método con su función. Permite escribir
# METODOS["Mínima varianza"](r), que equivale a minima_varianza(r).
METODOS = {"Máximo Sharpe": maximo_sharpe,
           "Mínima varianza": minima_varianza,
           "Pesos iguales": pesos_iguales}


# ----------------------------------------------------------------------------
# 4. Medidas de riesgo
# ----------------------------------------------------------------------------
def serie_portafolio(w, activos, r=LOG):
    """Rendimiento diario del portafolio: suma ponderada de rendimientos logarítmicos."""
    # r[activos].values : tabla de rendimientos solo de los activos del perfil, como matriz (días × activos).
    # @ w              : multiplica cada día por los pesos y suma → un rendimiento por día para todo el portafolio.
    # Ejemplo de un día: KO −1 % con peso 74 % y HD +2 % con peso 26 % → 0,74 × (−1 %) + 0,26 × 2 % = −0,22 %.
    # Si no se indica otra tabla, r=LOG usa los rendimientos de 2015-2024.
    return pd.Series(r[activos].values @ w, index=r.index)


def var_historico(r, confianza):
    # VaR (Valor en Riesgo, Value at Risk) histórico y CVaR (Valor en Riesgo Condicional, Conditional Value at Risk).
    # No supone ninguna distribución: trabaja directamente con lo que ocurrió en la historia.
    q = np.quantile(r, 1 - confianza)      # Percentil (1 − confianza). Con confianza 0,95 es el percentil 5: el
                                           # rendimiento que deja por debajo al 5 % de los peores días.
                                           # Ejemplo: de 2 500 días, separa los 125 peores del resto.
    return -q, -r[r <= q].mean()           # VaR = ese percentil con signo positivo (pérdida como número positivo).
                                           # CVaR = promedio de los días que quedaron por debajo del percentil, es decir,
                                           # cuánto se pierde EN PROMEDIO en esos días excepcionalmente desfavorables.


def var_parametrico(r, confianza, dias=1):
    # VaR paramétrico: supone que los rendimientos siguen una distribución normal (la campana de Gauss).
    mu, s = r.mean(), r.std(ddof=1)        # mu = rendimiento promedio diario; s = desviación estándar diaria
                                           # (ddof=1 la calcula como muestra, dividiendo entre n − 1).
    return -(mu * dias + norm.ppf(1 - confianza) * s * np.sqrt(dias))
                                           # norm.ppf(0,01) = −2,326: el valor z que deja el 1 % de la campana a su izquierda.
                                           # Para varios días, el promedio crece con los días y el riesgo con la raíz de
                                           # los días (regla de la raíz del tiempo). Ejemplo: a 10 días, s × √10 = s × 3,16.


def maxima_caida(valor):
    # Máxima caída (maximum drawdown): la mayor baja desde un máximo anterior hasta un mínimo posterior.
    # valor es la serie del valor de la cartera en el tiempo (por ejemplo, 1,00 → 1,05 → 0,90 → 1,10).
    pico = valor.cummax()                  # Máximo acumulado hasta cada día: 1,00 → 1,05 → 1,05 → 1,10.
    caida = valor / pico - 1               # Distancia de cada día respecto de su máximo: 0 % → 0 % → −14,3 % → 0 %.
    fondo = caida.idxmin()                 # Fecha del punto más bajo (la peor caída).
    inicio = valor.loc[:fondo].idxmax()    # Fecha del máximo previo a ese fondo: donde empezó la caída.
    rec = valor.loc[fondo:][valor.loc[fondo:] >= valor.loc[inicio]]
                                           # Días posteriores al fondo en que el valor volvió a alcanzar el máximo previo.
    dias = int(np.busday_count(fondo.date(), rec.index[0].date())) if len(rec) else None
                                           # Días hábiles entre el fondo y la primera recuperación. Si la cartera
                                           # todavía no se recuperó al cierre de 2024, queda en None (sin dato).
    return float(caida.min()), dias        # Devuelve la peor caída (número negativo, por ejemplo −0,25 = −25 %) y los días.


def resumen(r):
    """Rendimiento, volatilidad, Sharpe y máxima caída de una serie de rendimientos log."""
    rend = r.mean() * SESIONES             # Rendimiento anual = promedio diario × 252.
    vol = r.std() * np.sqrt(SESIONES)      # Volatilidad anual = desviación diaria × √252 (regla de la raíz del tiempo).
    mdd, dias = maxima_caida(np.exp(r.cumsum()))
                                           # r.cumsum() suma los rendimientos logarítmicos día a día; np.exp() los
                                           # convierte en el valor de 1 dólar invertido. Sobre esa curva se mide la caída.
    return {"rend": rend, "vol": vol, "sharpe": (rend - TASA_LIBRE) / vol,
            "mdd": mdd, "dias_rec": dias}  # Se devuelve un diccionario: cada resultado con su nombre.


# ----------------------------------------------------------------------------
# 5. Backtest con ventana móvil, rebalanceo mensual y costos
# ----------------------------------------------------------------------------
# El backtest (prueba retrospectiva) simula qué habría pasado si un gestor hubiera aplicado el método
# en tiempo real: cada mes recalcula los pesos con el año anterior, reajusta la cartera, paga la
# comisión y deja que el mercado la mueva hasta el mes siguiente. Nunca usa datos del futuro.

@lru_cache(maxsize=32)                     # Memoriza hasta 32 backtests distintos. El servidor lo pide varias veces con
                                           # los mismos datos (para el JSON y para los gráficos) y así responde al instante.
def backtest(activos, metodo):
    # activos llega como tupla, por ejemplo ("NFLX", "AMD", "HD", "KO", "BTC-USD"), porque lru_cache solo
    # puede memorizar datos que no cambian; una lista podría cambiar y no se admite.
    activos = list(activos)                # Se convierte en lista para poder seleccionar columnas con ella.
    fechas = SIMPLE.index                  # Todas las fechas de bolsa disponibles.
    meses = pd.Series(fechas, index=fechas).groupby([fechas.year, fechas.month]).first().values
                                           # Agrupa las fechas por año y mes, y toma la primera de cada grupo:
                                           # el primer día hábil de cada mes (02/01/2015, 02/02/2015, 02/03/2015, ...).
    meses = [m for m in meses if (fechas < m).sum() >= VENTANA]
                                           # Conserva solo los meses con al menos 252 días de historia previa.
                                           # Por eso el backtest empieza a inicios de 2016 y no en 2015.
    valor, w_act, costos = 1.0, None, 0.0  # La cartera arranca valiendo 1 dólar (después se multiplica por el monto real),
                                           # todavía sin pesos (None) y sin costos acumulados.
    valores, indice = [], []               # Listas vacías donde se irá anotando el valor de cada día y su fecha.
    for i, mes in enumerate(meses):        # Recorre los meses uno por uno; i es su número de orden (0, 1, 2, ...).
        pos = fechas.get_loc(mes)          # Posición de ese primer día del mes dentro de la lista de fechas.
        w_obj = METODOS[metodo](LOG[activos].iloc[pos - VENTANA:pos])
                                           # Pesos objetivo del mes, calculados con los 252 días ANTERIORES (de pos − 252
                                           # a pos − 1). El día de hoy no se incluye: el gestor aún no lo conoce.
        giro = 1.0 if w_act is None else np.abs(w_obj - w_act).sum()
                                           # Giro (rotación) = fracción de la cartera que hay que comprar o vender.
                                           # El primer mes se compra todo: giro = 1 (100 %).
                                           # Ejemplo después: pasar KO de 70 % a 75 % y HD de 30 % a 25 % → 5 % + 5 % = 10 %.
        pago = COSTO * giro * valor        # Comisión: 0,1 % × giro × valor de la cartera. En el ejemplo: 0,001 × 0,10 × valor.
        valor -= pago                      # La comisión se descuenta del valor de la cartera...
        costos += pago                     # ...y se acumula para informar el total pagado.
        w_act = w_obj.copy()               # Tras reajustar, la cartera queda con los pesos objetivo.
        fin = meses[i + 1] if i + 1 < len(meses) else None
                                           # Primer día del mes siguiente. En el último mes no hay siguiente (None).
        tramo = SIMPLE[activos].loc[mes:] if fin is None else SIMPLE[activos].loc[mes:fin].iloc[:-1]
                                           # Rendimientos simples de los días del mes: desde el primer día del mes hasta el
                                           # día anterior al mes siguiente (.iloc[:-1] quita ese primer día del mes siguiente,
                                           # que pertenece al próximo tramo). En el último mes, hasta el final de los datos.
        for fecha, fila in tramo.iterrows():          # Recorre el mes día por día.
            crece = w_act * (1 + fila.values)         # Cuánto vale ahora cada parte de la cartera. Ejemplo: 70 % en KO
                                                      # que sube 1 % → 0,70 × 1,01 = 0,707.
            valor *= crece.sum()                      # El valor total se multiplica por el crecimiento de todo el portafolio.
            w_act = crece / crece.sum()               # Los pesos «se mueven solos»: el activo que sube gana participación.
                                                      # Por eso cada mes hay que volver a reajustar.
            valores.append(valor)                     # Se anota el valor del día...
            indice.append(fecha)                      # ...y su fecha.
    return pd.Series(valores, index=indice), costos   # Devuelve la curva del valor (empieza cerca de 1) y el costo total.


def metricas_backtest(v):
    # Resume la curva del backtest v en cuatro indicadores.
    r = np.log(v / v.shift(1)).dropna()      # Rendimientos logarítmicos diarios de la curva.
    anos = len(r) / SESIONES                 # Duración en años: número de días / 252. Ejemplo: 2 263 días ≈ 9 años.
    mdd, _ = maxima_caida(v)                 # Máxima caída de la curva (los días de recuperación no se usan aquí: _).
    return {"crec": (v.iloc[-1] / v.iloc[0]) ** (1 / anos) - 1,
            # crec: tasa de crecimiento anual compuesto (CAGR, Compound Annual Growth Rate).
            # Ejemplo: si la cartera se multiplica por 3,08 en 9 años → 3,08^(1/9) − 1 ≈ 13,3 % al año.
            "sharpe": (r.mean() * SESIONES - TASA_LIBRE) / (r.std() * np.sqrt(SESIONES)),
            # sharpe: ratio de Sharpe de la curva, con rendimiento y volatilidad anualizados.
            "mdd": mdd, "multiplo": v.iloc[-1] / v.iloc[0]}
            # multiplo: cuántas veces se multiplicó el dinero (valor final / valor inicial).
            # El monto final en dólares se obtiene después: múltiplo × monto invertido.


# ----------------------------------------------------------------------------
# 6. Formato de cifras en español (1 234,5)
# ----------------------------------------------------------------------------
# Python escribe las cifras al estilo inglés (1,234.5). Estas tres funciones las pasan al formato
# que se usa en el informe: espacio para los miles y coma para los decimales.

def usd(x):
    # f"{x:,.0f}" escribe x sin decimales y con coma de miles: 73557.9 → "73,558".
    # .replace(",", " ") cambia la coma por un espacio: "73 558". Resultado final: "USD 73 558".
    return "USD " + f"{x:,.0f}".replace(",", " ")


def pct(x, dec=1):
    # Convierte una fracción en porcentaje con «dec» decimales (uno, si no se indica otra cosa).
    # Ejemplo: pct(0.73558) multiplica por 100 (73.558), redondea a un decimal ("73.6")
    # y cambia el punto por coma: "73,6 %".
    return f"{x * 100:.{dec}f}".replace(".", ",") + " %"


def num(x, dec=2):
    # Escribe un número con «dec» decimales y coma decimal. Ejemplo: num(0.8954) → "0,90".
    return f"{x:.{dec}f}".replace(".", ",")


# ----------------------------------------------------------------------------
# 7. Función principal: la que llama el servidor
# ----------------------------------------------------------------------------
# Recibe cuatro datos:
#   monto              : dólares a invertir. Ejemplo: 100000.
#   perfil             : "conservador", "moderado" o "agresivo".
#   empresa            : nombre del cliente; solo se repite en la respuesta. Ejemplo: "Walker y Wilson Company".
#   perdida_maxima_usd : pérdida diaria máxima que el cliente tolera, en dólares. Ejemplo: 2000.
#                        Es opcional: si no se envía, vale None y no se evalúa ningún límite.
# Devuelve un diccionario con todos los resultados (reparto, riesgo, prueba fuera de muestra y backtest).
def analizar(monto, perfil, empresa="", perdida_maxima_usd=None):
    perfil = perfil.strip().lower()        # Limpia el texto: quita espacios sobrantes y pasa a minúsculas.
                                           # Así " Conservador " se entiende igual que "conservador".
    if perfil not in PERFILES:             # Validación: el perfil debe ser uno de los tres conocidos.
        raise ValueError("El perfil debe ser conservador, moderado o agresivo.")
                                           # raise detiene la función y envía este mensaje; el servidor lo convierte en
                                           # una respuesta 400 (solicitud incorrecta) para el agente.
    if monto <= 0:                         # Validación: el monto debe ser positivo.
        raise ValueError("El monto debe ser mayor que cero.")

    metodo = PERFILES[perfil]["metodo"]    # Método del perfil. Conservador → "Mínima varianza".
    activos = PERFILES[perfil]["activos"]  # Activos del perfil. Conservador → las cuatro acciones y el bitcoin.

    # 7.1 Reparto recomendado (estimado con 2015-2024)
    w = METODOS[metodo](LOG[activos])      # Pesos óptimos con TODA la historia 2015-2024 de los activos del perfil.
    rend, vol, sharpe = estadisticas(w, LOG[activos])
                                           # Rendimiento esperado, volatilidad y Sharpe anuales de esos pesos.
    orden = np.argsort(-w)                 # Posiciones de los activos ordenadas de mayor a menor peso
                                           # (argsort ordena de menor a mayor; el signo − invierte el orden).
    reparto = "; ".join(f"{activos[i]}: {pct(w[i])} = {usd(w[i] * monto)}"
                        for i in orden if w[i] >= 0.0005)
                                           # Arma el texto del reparto, activo por activo, separado por «; ».
                                           # Ejemplo: "KO: 73,6 % = USD 73 558; HD: 17,1 % = USD 17 051; ...".
                                           # Solo muestra los activos con peso de al menos 0,05 %: un activo con peso
                                           # prácticamente nulo (como AMD en el perfil conservador) no aparece.

    # 7.2 Riesgo del portafolio recomendado
    r = serie_portafolio(w, activos)       # Rendimiento diario del portafolio en cada día de 2015-2024.
    var95, cvar95 = var_historico(r.values, 0.95)
                                           # VaR (Valor en Riesgo) y CVaR (Valor en Riesgo Condicional) al 95 % a un día,
                                           # como fracción. Ejemplo: var95 = 0,01478 → 1,478 % del monto.
    var99, _ = var_historico(r.values, 0.99)
                                           # VaR al 99 % a un día: el límite que solo se supera 1 de cada 100 días.
    var99_10 = var_parametrico(r, 0.99, 10)
                                           # VaR paramétrico al 99 % a 10 días: la medida que usan los bancos según
                                           # las normas de Basilea.
    mdd, dias_rec = maxima_caida(np.exp(r.cumsum()))
                                           # Máxima caída del portafolio y días hábiles que tardó en recuperarse.

    # 7.3 Prueba fuera de muestra: pesos de 2015-2019 aplicados a 2020-2024
    # Pregunta que responde: «si en 2019 hubiéramos elegido los pesos con lo que sabíamos entonces,
    # ¿se habría cumplido lo que prometían en 2020-2024?». Se compara el Sharpe prometido con el logrado.
    dentro = LOG.loc[:CORTE_MUESTRA]       # Muestra de estimación: 2015-2019 (dentro de la muestra).
    fuera = LOG.loc[CORTE_MUESTRA:].iloc[1:]
                                           # Muestra de evaluación: desde el 31/12/2019 en adelante; .iloc[1:] quita ese
                                           # primer día, que ya pertenece a la muestra de estimación. Queda 2020-2024.
    comparacion = []                       # Lista donde se acumula una frase por cada método.
    for nombre in METODOS:                 # Se prueban los TRES métodos, no solo el del perfil, para compararlos.
        cols = ACCIONES if nombre == "Pesos iguales" else ACCIONES + CRIPTO
                                           # Pesos iguales usa solo acciones; los otros dos incluyen el bitcoin.
        w_d = METODOS[nombre](dentro[cols])                 # Pesos estimados con 2015-2019.
        promesa = estadisticas(w_d, dentro[cols])[2]        # Sharpe que esos pesos «prometían» en 2015-2019.
        real = resumen(serie_portafolio(w_d, cols, fuera))  # Lo que esos mismos pesos lograron en 2020-2024.
        etiqueta = nombre + (" (sin Bitcoin)" if cols == ACCIONES else " (con Bitcoin)")
                                           # Nombre del método con la aclaración de si incluye bitcoin.
        comparacion.append(f"{etiqueta}: Sharpe prometido {num(promesa)} y logrado {num(real['sharpe'])}, "
                           f"rendimiento {pct(real['rend'])} al año, máxima caída {pct(real['mdd'])}")
                                           # Frase de resultado del método, con las cifras en formato español.
        if nombre == metodo:               # Si es el método del perfil del cliente...
            fm = (promesa, real)           # ...se guardan sus cifras aparte para la respuesta.
    sp_fm = resumen(fuera[INDICE])         # El S&P 500 en 2020-2024, como referencia de mercado.
    comparacion.append(f"S&P 500: Sharpe {num(sp_fm['sharpe'])}, rendimiento {pct(sp_fm['rend'])} al año, "
                       f"máxima caída {pct(sp_fm['mdd'])}")

    # 7.4 Backtest 2016-2024 del método elegido, de los pesos iguales y del S&P 500
    v, costos = backtest(tuple(activos), metodo)
                                           # Curva del método del perfil y comisiones totales (como fracción del monto).
    bt = metricas_backtest(v)              # Crecimiento anual, Sharpe, máxima caída y múltiplo de esa curva.
    v_ig, _ = backtest(tuple(ACCIONES), "Pesos iguales")
                                           # Curva de referencia: las cuatro acciones a partes iguales.
    sp = PRECIOS[INDICE].loc[v.index[0]:]  # Precios del S&P 500 desde el mismo día en que empieza el backtest.
    sp_bt = metricas_backtest(sp / sp.iloc[0])
                                           # Se divide entre el primer precio para que el índice también arranque en 1.

    # 7.5 Límite de pérdida fijado por la empresa
    if perdida_maxima_usd and perdida_maxima_usd > 0:     # Solo si el cliente indicó un límite positivo.
        cumple = bool(var95 * monto <= perdida_maxima_usd)
                                           # True si el VaR al 95 % en dólares no pasa del límite.
                                           # Ejemplo: 0,01478 × 100 000 = 1 478 ≤ 2 000 → True (cumple).
        limite = (f"El VaR al 95 % ({usd(var95 * monto)}) {'NO supera' if cumple else 'SUPERA'} "
                  f"la pérdida máxima tolerada de {usd(perdida_maxima_usd)}.")
                                           # Frase lista para el informe con el veredicto.
    else:
        cumple, limite = None, "La empresa no fijó una pérdida máxima tolerada."

    # Diccionario de salida: es exactamente lo que el servidor envía al agente en formato JSON.
    # Las cifras «_pct» van en porcentaje (ej. 13.3) y las «_usd» en dólares redondeados al entero.
    # round(x, n) redondea x a n decimales.
    salida = {
        "empresa": empresa,                                          # Nombre del cliente.
        "monto_usd": round(monto, 2),                                # Monto invertido.
        "perfil": perfil,                                            # Perfil ya limpio.
        "metodo": metodo,                                            # Método de optimización aplicado.
        "activos": ", ".join(activos),                               # Activos permitidos, separados por coma.
        "reparto": reparto,                                          # Texto del reparto (7.1).
        "rendimiento_esperado_anual_pct": round(rend * 100, 2),      # Rendimiento anual esperado 2015-2024.
        "volatilidad_anual_pct": round(vol * 100, 2),                # Volatilidad anual.
        "ratio_sharpe": round(sharpe, 3),                            # Ratio de Sharpe.
        "var95_1dia_usd": round(var95 * monto, 0),                   # VaR al 95 % a un día, en dólares.
        "var99_1dia_usd": round(var99 * monto, 0),                   # VaR al 99 % a un día, en dólares.
        "cvar95_1dia_usd": round(cvar95 * monto, 0),                 # CVaR al 95 % a un día, en dólares.
        "var99_10dias_usd": round(var99_10 * monto, 0),              # VaR paramétrico al 99 % a 10 días, en dólares.
        "maxima_caida_pct": round(mdd * 100, 1),                     # Máxima caída 2015-2024 (número negativo).
        "dias_recuperacion": dias_rec,                               # Días hábiles para recuperarse de esa caída.
        "cumple_limite_perdida": cumple,                             # True, False o None (sin límite) (7.5).
        "limite_perdida": limite,                                    # Frase del veredicto (7.5).
        "fuera_muestra_sharpe_prometido": round(fm[0], 3),           # Sharpe que prometía el método en 2015-2019.
        "fuera_muestra_sharpe_logrado": round(fm[1]["sharpe"], 3),   # Sharpe que logró en 2020-2024.
        "fuera_muestra_rendimiento_pct": round(fm[1]["rend"] * 100, 2),        # Rendimiento anual logrado en 2020-2024.
        "sp500_fuera_muestra_rendimiento_pct": round(sp_fm["rend"] * 100, 2),  # El del S&P 500 en el mismo periodo.
        "comparacion_metodos": " | ".join(comparacion),              # Las frases de los tres métodos y del S&P 500.
        "backtest_periodo": f"{v.index[0]:%d/%m/%Y} al {v.index[-1]:%d/%m/%Y}",  # Primer y último día del backtest.
        "backtest_valor_final_usd": round(bt["multiplo"] * monto, 0),          # En qué se convirtió el monto.
        "backtest_crecimiento_anual_pct": round(bt["crec"] * 100, 2),          # Crecimiento anual compuesto.
        "backtest_sharpe": round(bt["sharpe"], 3),                             # Sharpe del backtest.
        "backtest_maxima_caida_pct": round(bt["mdd"] * 100, 1),                # Peor caída del backtest.
        "backtest_costos_usd": round(costos * monto, 0),                       # Comisiones pagadas en total.
        "backtest_pesos_iguales_valor_final_usd": round(v_ig.iloc[-1] / v_ig.iloc[0] * monto, 0),  # Referencia 1.
        "backtest_sp500_valor_final_usd": round(sp_bt["multiplo"] * monto, 0),                     # Referencia 2.
        "datos": f"Yahoo Finance, {PRECIOS.index[0]:%d/%m/%Y} al {PRECIOS.index[-1]:%d/%m/%Y}, "
                 f"{len(LOG)} sesiones; tasa libre de riesgo {pct(TASA_LIBRE, 0)}",  # Ficha de los datos usados.
        "version_modelo": VERSION,                                   # Versión del modelo que hizo el cálculo.
    }
    # números de numpy -> números de Python, para que viajen en JSON
    # Algunas cifras salen como tipos de NumPy (por ejemplo, numpy.float64), que el formato JSON no reconoce.
    # .item() las convierte en números normales de Python; lo que no tiene .item() (los textos) pasa tal cual.
    return {k: (v.item() if hasattr(v, "item") else v) for k, v in salida.items()}


# Este bloque solo se ejecuta cuando se corre el archivo directamente («python modelo.py»);
# no se ejecuta cuando el servidor lo importa. Sirve para probar el modelo por su cuenta:
# imprime el resultado completo de los tres perfiles. 100_000 es otra forma de escribir 100000
# (el guion bajo solo facilita la lectura).
if __name__ == "__main__":
    import json
    for p in PERFILES:
        print(json.dumps(analizar(100_000, p, "Walker y Wilson", 2_000), ensure_ascii=False, indent=1))
                                           # json.dumps convierte el diccionario en texto JSON; ensure_ascii=False
                                           # conserva las tildes e indent=1 lo muestra con sangría, fácil de leer.
