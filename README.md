# Curva S y avance de obra por partidas

Este ejercicio genera un cronograma sintético de una obra con 12 partidas y 24 periodos quincenales. El presupuesto de cada partida determina su peso en el avance físico valorizado.

## Archivos principales

- `generar_curva_s.py`: genera los datos, crea o lee el Excel y produce los gráficos.
- `datos_avance_obra.xlsx`: contiene el resumen, las partidas, el avance periódico y el detalle por partida.
- `curva_s.png`: compara el avance acumulado planificado y real.
- `avance_por_partidas.png`: compara el avance planificado y real de cada partida a la fecha de corte.

## Ejecución

Desde la carpeta raíz del proyecto, active el entorno virtual y ejecute:

```powershell
.\.venv\Scripts\Activate.ps1
python .\outputs\generar_curva_s.py
```

Para volver a crear únicamente los gráficos a partir del Excel existente:

```powershell
python .\outputs\generar_curva_s.py --solo-graficos
```

## Panel web local

El panel Streamlit muestra las métricas, los gráficos y las tablas del Excel. Ejecútelo desde la raíz del proyecto:

```powershell
.\.venv\Scripts\Activate.ps1
streamlit run .\outputs\app_streamlit.py
```

Para desplegarla en Streamlit Community Cloud, suba `app_streamlit.py`, `generar_curva_s.py`, `requirements.txt` y este archivo README a la raíz de un repositorio de GitHub. El panel generará el Excel y las imágenes por sí solo en el primer arranque.

## Lógica del código

1. Define partidas con presupuesto, periodo de inicio, duración planificada y retraso realista.
2. Distribuye el avance de cada partida con una curva de producción tipo campana.
3. Pondera cada avance por el peso presupuestal de la partida.
4. Acumula los aportes para obtener la curva S del proyecto.
5. Guarda las tablas en Excel y las vuelve a leer con `pandas`.
6. Crea los gráficos con `matplotlib`, usando fondo negro y una paleta morada.

Los datos son sintéticos y reproducibles. Puede cambiar `PARTIDAS`, `PERIODO_CORTE`, `FECHA_INICIO` o la semilla para ensayar otros escenarios.
