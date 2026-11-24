# Asesor de Portafolios Inteligentes · modelo del TAF · v3.0

Curso 012755 Construcción de Portafolios Inteligentes · CENTRUM PUCP

Servidor que recibe del agente de Copilot Studio el monto, el perfil de riesgo y la
pérdida máxima tolerada, y devuelve el reparto del portafolio, su riesgo (VaR, CVaR y
máxima caída), la prueba fuera de muestra y el backtest frente al S&P 500. Desde la
versión 1.1 también guarda el informe que redacta el agente y entrega un enlace para
descargarlo como archivo .md; desde la versión 1.2 entrega además un segundo enlace con el
informe en Word (.docx).

| Archivo | Qué contiene |
|---|---|
| modelo.py | Todo el cálculo financiero |
| app.py | El servidor (FastAPI): rutas /health, /analizar, /informe, /descargar y /descargar/{id}/word |
| precios.csv | Precios de cierre ajustados de Yahoo Finance, 2015-2024 |
| graficos.py | Los seis gráficos del informe en Word, al estilo TradingView |
| requirements.txt | Bibliotecas que instala Railway (incluye python-docx y matplotlib para el Word) |
| Procfile | Comando con que Railway arranca el servidor |

Prueba local: `uvicorn app:app --reload` y abrir http://127.0.0.1:8000/docs

## Historial de versiones

| Versión | Cambio |
|---|---|
| 1.0 | Servidor con /health y /analizar |
| 1.1 | Rutas /informe y /descargar/{id}: el agente guarda su informe y entrega el enlace de descarga en .md |
| 1.2 | Ruta /descargar/{id}/word: el mismo informe se descarga también en Word (.docx); /informe devuelve los dos enlaces |
| 1.3 | El Word trae seis gráficos al estilo TradingView: reparto, velas del S&P 500, base 100, ganancias y pérdidas diarias con el VaR, backtest y caídas |
| 1.31 | Código comentado línea por línea en app.py, modelo.py (1.1) y graficos.py (1.1); el funcionamiento y las cifras no cambian |
| 3.0 | Numeración unificada del servidor, modelo.py y graficos.py con el material del curso; el funcionamiento y las cifras no cambian |
