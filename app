import pandas as pd
import streamlit as st

import comparador as cmp

st.set_page_config(page_title="Comparador Estructura Programática", page_icon="📊", layout="wide")

st.markdown(
    """
    <style>
    h1, h2, h3 { color: #9F2241; }
    div[data-testid="stMetricValue"] { color: #9F2241; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Comparador de Estructura Programática")
st.caption("Compara la base del año nuevo contra la del año anterior con la llave "
           "UR-FI-FN-SF-AI-Mod+Pp-Partida (UR y partida amarradas) y marca lo que nunca se había usado.")

c1, c2 = st.columns(2)
with c1:
    anio_n = st.text_input("Año nuevo", "2027")
    f_n = st.file_uploader(f"Base {anio_n}", type=["xlsx", "xlsm", "xls", "csv", "txt"], key="n")
with c2:
    anio_a = st.text_input("Año anterior", "2026")
    f_a = st.file_uploader(f"Base {anio_a}", type=["xlsx", "xlsm", "xls", "csv", "txt"], key="a")


def configurar(archivo, etiqueta, key):
    """Elige hoja, detecta columnas y permite corregir el mapeo."""
    hojas = cmp.hojas_de(archivo)
    # Preferir la hoja con datos planos (no tablas dinámicas)
    pref = next((i for i, h in enumerate(hojas) if not h.upper().startswith("TD")), 0)
    hoja = st.selectbox(f"Hoja de la base {etiqueta}", hojas, index=pref, key=f"h{key}") if len(hojas) > 1 else None
    df = cmp.leer_base(archivo, None if hoja == "(CSV)" else hoja)
    mapeo = cmp.detectar_columnas(df)

    with st.expander(f"Columnas detectadas — base {etiqueta} ({len(df):,} filas)",
                     expanded=any(c not in mapeo for c in cmp.CAMPOS if c != "PARTIDA") or
                     ("PARTIDA" not in mapeo and "_PARTES" not in mapeo)):
        opciones = ["—"] + list(df.columns)
        cols = st.columns(5)
        for i, campo in enumerate(cmp.CAMPOS):
            if campo == "PARTIDA" and "_PARTES" in mapeo:
                cols[i % 5].markdown("**Partida**: se arma con Capítulo+Concepto+Genérica+Específica")
                continue
            actual = mapeo.get(campo)
            elegido = cols[i % 5].selectbox(cmp.ETIQUETAS[campo], opciones,
                                            index=opciones.index(actual) if actual in opciones else 0,
                                            key=f"{key}{campo}")
            if elegido != "—":
                mapeo[campo] = elegido
            else:
                mapeo.pop(campo, None)
        st.dataframe(df.head(5), use_container_width=True)
    faltan = [cmp.ETIQUETAS[c] for c in cmp.LLAVE
              if c not in mapeo and c != "MOD" and not (c == "PARTIDA" and "_PARTES" in mapeo)]
    return df, mapeo, faltan


if f_n and f_a:
    with c1:
        df_n, map_n, falt_n = configurar(f_n, anio_n, "n")
    with c2:
        df_a, map_a, falt_a = configurar(f_a, anio_a, "a")

    if falt_n or falt_a:
        st.error(f"Faltan columnas por asignar — {anio_n}: {falt_n or 'ninguna'} | {anio_a}: {falt_a or 'ninguna'}")
        st.stop()

    n27 = cmp.normalizar(df_n, map_n)
    n26 = cmp.normalizar(df_a, map_a)
    r = cmp.comparar(n27, n26)
    comp = r["comparativo"]
    nuevas = comp[comp["ESTATUS"].str.startswith("NUEVA")]

    st.divider()
    m = st.columns(4)
    m[0].metric(f"Llaves {anio_n}", f"{len(comp):,}")
    m[1].metric("Nuevas (nunca usadas)", f"{len(nuevas):,}", f"{len(nuevas) / max(len(comp), 1):.1%}",
                delta_color="off")
    m[2].metric("Importe en llaves nuevas", f"${nuevas['IMPORTE_2027'].sum():,.0f}")
    m[3].metric(f"Llaves solo en {anio_a}", f"{len(r['solo_2026']):,}")

    t1, t2, t3, t4 = st.tabs(["Resumen por UR", "Nuevas", "Comparativo completo", f"Solo {anio_a}"])
    with t1:
        st.dataframe(r["resumen"].rename(columns={
            "LLAVES_2027": f"Llaves {anio_n}", "NUEVAS": "Nuevas", "IMPORTE_2027": f"Importe {anio_n}",
            "IMPORTE_NUEVAS": "Importe nuevas", "SOLO_2026": f"Solo {anio_a}"}),
            use_container_width=True, hide_index=True)
        st.dataframe(nuevas["MOTIVO"].value_counts().sort_index().rename("Llaves").to_frame(),
                     use_container_width=True)
    with t2:
        urs = st.multiselect("Filtrar UR", sorted(nuevas["UR"].unique()))
        mot = st.multiselect("Filtrar motivo", sorted(nuevas["MOTIVO"].unique()))
        v = nuevas
        if urs:
            v = v[v["UR"].isin(urs)]
        if mot:
            v = v[v["MOTIVO"].isin(mot)]
        st.dataframe(v[["UR", "FI", "FN", "SF", "AI", "MOD", "PP", "PARTIDA", "IMPORTE_2027",
                        "MOTIVO", "UR_PARTIDA_2026"]], use_container_width=True, hide_index=True)
    with t3:
        st.dataframe(comp.drop(columns=["LLAVE_ESTRUCTURA", "UR_PARTIDA", "UR_PP"]),
                     use_container_width=True, hide_index=True)
    with t4:
        st.dataframe(r["solo_2026"].drop(columns=["LLAVE_ESTRUCTURA", "UR_PARTIDA", "UR_PP"]),
                     use_container_width=True, hide_index=True)

    xls = cmp.generar_excel(r, anio_n, anio_a)
    st.download_button("⬇️ Descargar Excel comparativo", xls,
                       file_name=f"Comparativo_Estructura_{anio_n}_vs_{anio_a}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")
else:
    st.info("Sube ambas bases (Excel o CSV). La app detecta solas las columnas de UR, estructura, "
            "partida e importe; si alguna no se reconoce, puedes asignarla a mano.")
