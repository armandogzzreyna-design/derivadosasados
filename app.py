from __future__ import annotations

from datetime import date
from pathlib import Path

import streamlit as st

from src.excel_utils import fetch_fix_value, make_workdir, save_upload, save_uploads, workbook_bytes
from src.listados import run_listados_process
from src.otc import run_otc_process
from src.pdf_statements import last_business_day, parse_statement_uploads


st.set_page_config(page_title="Validador Derivados", page_icon="📊", layout="wide")


def upload_excel(label: str, key: str, required: bool = False):
    help_text = "Obligatorio" if required else "Opcional"
    return st.file_uploader(label, type=["xlsx", "xlsm"], key=key, help=help_text)


def show_result(result, file_name: str):
    for msg in result.messages:
        st.success(msg)
    for warning in result.warnings:
        st.warning(warning)
    st.download_button(
        "Descargar resultado",
        data=workbook_bytes(result.path),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


st.title("Validador de Derivados")
st.caption("Listados y OTC en una sola app, con cargas separadas para que cada archivo caiga donde corresponde.")

with st.sidebar:
    st.header("Tipo de cambio FIX")
    auto_fix = st.toggle("Consultar Banxico", value=False)
    manual_fix = st.number_input("FIX manual", min_value=0.0, value=17.0000, step=0.0001, format="%.4f")
    if auto_fix and st.button("Obtener FIX", use_container_width=True):
        try:
            st.session_state["fix_value"] = fetch_fix_value()
            st.success(f"FIX: {st.session_state['fix_value']:.4f}")
        except Exception as exc:
            st.error(f"No se pudo consultar Banxico: {exc}")
    fix_value = float(st.session_state.get("fix_value", manual_fix))
    st.divider()
    st.write(f"FIX a usar: **{fix_value:.4f}**")

tab_listados, tab_otc = st.tabs(["Listados", "OTC / CSA"])

with tab_listados:
    st.subheader("Carga de archivos Listados")
    col_a, col_b = st.columns(2)
    with col_a:
        listados_template = upload_excel("Plantilla Validacion Derivados Listados", "listados_template", required=True)
        vector_file = upload_excel("LISTADOS / Vector.xlsx", "vector_file")
        llamadas_file = upload_excel("LISTADOS / Llamadas_retiros.xlsx", "llamadas_file")
        int_com_file = upload_excel("LISTADOS / INT & COM.xlsx", "int_com_file")
    with col_b:
        vectores_file_listados = upload_excel("VECTORES.xlsx", "vectores_file_listados")
        aims_file = upload_excel("ResultadosDetalleAIM.xlsx", "aims_file")
        valuation_date = st.date_input("Fecha de valuacion para PDFs", value=last_business_day(date.today()), key="valuation_date")

    with st.expander("Estados de cuenta PDF"):
        pdf_cols = st.columns(2)
        with pdf_cols[0]:
            goldman_pdfs = st.file_uploader("GOLDMAN CME", type=["pdf"], accept_multiple_files=True, key="goldman_pdfs")
            santander_cme_pdfs = st.file_uploader("SANTANDER CME", type=["pdf"], accept_multiple_files=True, key="santander_cme_pdfs")
        with pdf_cols[1]:
            santander_mexder_pdfs = st.file_uploader("SANTANDER MEXDER", type=["pdf"], accept_multiple_files=True, key="santander_mexder_pdfs")
            scotia_mexder_pdfs = st.file_uploader("SCOTIA MEXDER", type=["pdf"], accept_multiple_files=True, key="scotia_mexder_pdfs")

    if st.button("Procesar Listados", type="primary", use_container_width=True):
        if listados_template is None:
            st.error("Carga la plantilla de Listados para continuar.")
        else:
            workdir = make_workdir()
            paths = workdir / "listados"
            path_listados = save_upload(listados_template, paths, "Validacion Derivados Listados.xlsx")
            path_vector = save_upload(vector_file, paths) if vector_file else None
            path_llamadas = save_upload(llamadas_file, paths) if llamadas_file else None
            path_int_com = save_upload(int_com_file, paths) if int_com_file else None
            path_vectores = save_upload(vectores_file_listados, paths) if vectores_file_listados else None
            path_aims = save_upload(aims_file, paths) if aims_file else None
            pdf_dir = paths / "pdfs"
            estados_df = parse_statement_uploads(
                goldman_cme=save_uploads(goldman_pdfs, pdf_dir / "goldman"),
                santander_cme=save_uploads(santander_cme_pdfs, pdf_dir / "santander_cme"),
                santander_mexder=save_uploads(santander_mexder_pdfs, pdf_dir / "santander_mexder"),
                scotia_mexder=save_uploads(scotia_mexder_pdfs, pdf_dir / "scotia_mexder"),
                valuation_date=valuation_date,
            )
            with st.spinner("Procesando Listados..."):
                result = run_listados_process(
                    Path(path_listados),
                    fix_value=fix_value,
                    path_vector=path_vector,
                    path_llamadas=path_llamadas,
                    path_int_com=path_int_com,
                    path_vectores=path_vectores,
                    path_aims=path_aims,
                    estados_df=estados_df,
                )
            show_result(result, "Validacion Derivados Listados procesado.xlsx")
            if not estados_df.empty:
                st.dataframe(estados_df, use_container_width=True)

with tab_otc:
    st.subheader("Carga de archivos OTC / CSA")
    col_a, col_b = st.columns(2)
    with col_a:
        otc_template = upload_excel("Plantilla Validaciones OTC CSA", "otc_template", required=True)
        vectores_file_otc = upload_excel("VECTORES.xlsx", "vectores_file_otc")
        deuda_file = st.file_uploader("DEUDA.076", type=["076", "txt"], key="deuda_file")
        derivados_file = st.file_uploader("DERIVADOS.077", type=["077", "txt"], key="derivados_file")
    with col_b:
        movimientos_otc = upload_excel("OTC / Movimientos_OTC.xlsx", "movimientos_otc")
        movimientos_cash = upload_excel("OTC / Movimientos_CASH.xlsx", "movimientos_cash")

    if st.button("Procesar OTC / CSA", type="primary", use_container_width=True):
        if otc_template is None:
            st.error("Carga la plantilla OTC / CSA para continuar.")
        else:
            workdir = make_workdir()
            paths = workdir / "otc"
            path_otc = save_upload(otc_template, paths, "Validaciones OTC CSA.xlsx")
            path_vectores = save_upload(vectores_file_otc, paths) if vectores_file_otc else None
            path_deuda = save_upload(deuda_file, paths) if deuda_file else None
            path_derivados = save_upload(derivados_file, paths) if derivados_file else None
            path_mov_otc = save_upload(movimientos_otc, paths) if movimientos_otc else None
            path_mov_cash = save_upload(movimientos_cash, paths) if movimientos_cash else None
            with st.spinner("Procesando OTC / CSA..."):
                result = run_otc_process(
                    Path(path_otc),
                    fix_value=fix_value,
                    vectores_path=path_vectores,
                    deuda_path=path_deuda,
                    derivados_path=path_derivados,
                    movimientos_otc=path_mov_otc,
                    movimientos_cash=path_mov_cash,
                )
            show_result(result, "Validaciones OTC CSA procesado.xlsx")
