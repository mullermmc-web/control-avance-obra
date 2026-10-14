"""Panel web local para consultar el avance de obra generado en el ejercicio."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st


BASE_DIR = Path(__file__).resolve().parent
RUTA_EXCEL = BASE_DIR / "datos_avance_obra.xlsx"
RUTA_CURVA_S = BASE_DIR / "curva_s.png"
RUTA_PARTIDAS = BASE_DIR / "avance_por_partidas.png"


def leer_hoja_flexible(ruta_excel: Path, hoja: str, columna_clave: str) -> pd.DataFrame:
    """Localiza la fila de encabezados tanto en Exceles simples como formateados."""
    muestra = pd.read_excel(ruta_excel, sheet_name=hoja, header=None, nrows=8)
    fila_cabecera = next(
        (
            indice
            for indice, fila in muestra.iterrows()
            if columna_clave in fila.astype(str).tolist()
        ),
        None,
    )
    if fila_cabecera is None:
        raise ValueError(f"No se encontró '{columna_clave}' en la hoja '{hoja}'.")
    return pd.read_excel(ruta_excel, sheet_name=hoja, header=fila_cabecera)


def asegurar_archivos() -> None:
    """Genera los insumos locales cuando la app arranca en un servidor nuevo."""
    if not RUTA_EXCEL.exists():
        from generar_curva_s import exportar_excel, generar_datos_sinteticos

        partidas, avance, detalle = generar_datos_sinteticos()
        exportar_excel(RUTA_EXCEL, partidas, avance, detalle)

    if not RUTA_CURVA_S.exists() or not RUTA_PARTIDAS.exists():
        from generar_curva_s import graficar_avance_partidas, graficar_curva_s, leer_excel

        partidas, avance = leer_excel(RUTA_EXCEL)
        graficar_curva_s(avance, RUTA_CURVA_S)
        graficar_avance_partidas(partidas, RUTA_PARTIDAS)


@st.cache_data
def cargar_datos() -> tuple[pd.DataFrame, pd.DataFrame]:
    partidas = leer_hoja_flexible(RUTA_EXCEL, "Partidas", "Partida")
    avance = leer_hoja_flexible(RUTA_EXCEL, "Avance_periodico", "Fecha")
    avance["Fecha"] = pd.to_datetime(avance["Fecha"])
    return partidas, avance


def configurar_estilo() -> None:
    st.markdown(
        """
        <style>
            .stApp { background: #09070d; color: #f6f0fa; }
            [data-testid="stHeader"] { background: rgba(9, 7, 13, 0.92); }
            [data-testid="stMetric"] {
                background: #15111b;
                border: 1px solid #5a189a;
                border-radius: 12px;
                padding: 16px;
            }
            [data-testid="stMetricLabel"], [data-testid="stMetricValue"] { color: #f6f0fa; }
            div[data-baseweb="tab-list"] { gap: 18px; }
            button[data-baseweb="tab"] { color: #c9b6d4; }
            button[data-baseweb="tab"][aria-selected="true"] { color: #e0aaff; border-bottom-color: #9d4edd; }
            .stDataFrame { border: 1px solid #5a189a; border-radius: 8px; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(page_title="Control de avance de obra", page_icon="📈", layout="wide")
    configurar_estilo()

    st.title("Control de avance de obra")
    st.caption("Panel local alimentado por el archivo Excel de datos sintéticos.")

    try:
        asegurar_archivos()
        partidas, avance = cargar_datos()
    except Exception as error:
        st.error(f"No fue posible leer el Excel: {error}")
        st.stop()

    avance_real = avance.dropna(subset=["Avance real acumulado"])
    ultimo = avance_real.iloc[-1]
    plan_corte = float(avance.loc[avance["Periodo"] == ultimo["Periodo"], "Avance plan acumulado"].iloc[0])
    real_corte = float(ultimo["Avance real acumulado"])
    desviacion = real_corte - plan_corte
    presupuesto = float(partidas["Presupuesto (S/)"].sum())

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Presupuesto total", f"S/ {presupuesto:,.0f}")
    col2.metric("Fecha de corte", f"{ultimo['Fecha']:%d/%m/%Y}")
    col3.metric("Avance real", f"{real_corte:.1%}", f"Plan {plan_corte:.1%}", delta_color="off")
    col4.metric("Desviación", f"{desviacion:.1%}", "Real − Plan", delta_color="inverse")

    tab_resumen, tab_partidas, tab_datos = st.tabs(["Gráficos", "Partidas", "Datos"])

    with tab_resumen:
        izquierda, derecha = st.columns(2)
        with izquierda:
            st.subheader("Curva S")
            if RUTA_CURVA_S.exists():
                st.image(str(RUTA_CURVA_S), use_container_width=True)
            else:
                st.info("El gráfico no existe todavía. Ejecute generar_curva_s.py.")
        with derecha:
            st.subheader("Avance por partidas")
            if RUTA_PARTIDAS.exists():
                st.image(str(RUTA_PARTIDAS), use_container_width=True)
            else:
                st.info("El gráfico no existe todavía. Ejecute generar_curva_s.py.")

    with tab_partidas:
        columnas = [
            "ID", "Partida", "Presupuesto (S/)", "Peso", "Avance plan al corte",
            "Avance real al corte", "Variación (pp)", "Estado",
        ]
        vista = partidas[columnas].copy()
        for columna in ["Peso", "Avance plan al corte", "Avance real al corte", "Variación (pp)"]:
            vista[columna] = vista[columna].map(lambda valor: f"{valor:.1%}")
        vista["Presupuesto (S/)"] = vista["Presupuesto (S/)"].map(lambda valor: f"S/ {valor:,.0f}")
        st.dataframe(vista, use_container_width=True, hide_index=True)

    with tab_datos:
        st.subheader("Avance quincenal")
        tabla_avance = avance.copy()
        tabla_avance["Fecha"] = tabla_avance["Fecha"].dt.strftime("%d/%m/%Y")
        st.dataframe(tabla_avance, use_container_width=True, hide_index=True)
        st.download_button(
            "Descargar Excel de datos",
            data=RUTA_EXCEL.read_bytes(),
            file_name=RUTA_EXCEL.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
