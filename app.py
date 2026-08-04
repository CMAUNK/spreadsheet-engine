"""SmartSheet v0.1 — preenchedor genérico de arquivos Excel."""
from __future__ import annotations

import base64
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from config_manager import load_template, load_units, save_template, save_units
from excel_pdf_converter import PdfConversionError, convert_workbook_to_pdf
from excel_manager import (
    build_filled_workbook,
    detect_header_candidates,
    find_matching_references,
    get_sheet_names,
    model_fingerprint,
    read_units,
    read_sheet_map,
    workbook_fingerprint,
)
from parser import parse_entries
from pdf_generator import generate_pdf
from pdf_preview import pdf_page_count, render_pdf_page
from utils import column_label, normalize_reference, safe_output_name


st.set_page_config(page_title="SmartSheet", page_icon="📄", layout="wide")


def initialize_state() -> None:
    defaults = {"parsed_rows": [], "parse_warnings": [], "source_key": None, "preview_zoom": 1.0, "generated_output": None}
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def reset_preview(source_key: str) -> None:
    if st.session_state.source_key != source_key:
        st.session_state.source_key = source_key
        st.session_state.parsed_rows = []
        st.session_state.parse_warnings = []
        st.session_state.generated_output = None
        st.session_state.preview_zoom = 1.0


def main() -> None:
    initialize_state()
    st.title("📄 SmartSheet")
    st.caption("Preencha qualquer modelo Excel sem alterar o arquivo original.")

    uploaded = st.file_uploader("1. Envie uma planilha Excel", type=["xlsx"])
    if not uploaded:
        st.info("Envie um arquivo .xlsx para começar.")
        return

    source_bytes = uploaded.getvalue()
    source_key = workbook_fingerprint(source_bytes)
    reset_preview(source_key)

    try:
        sheet_names = get_sheet_names(source_bytes)
    except Exception as exc:
        st.error(f"Não foi possível abrir este arquivo: {exc}")
        return

    remembered = load_template(source_key)
    default_sheet = remembered.get("sheet_name") if remembered else sheet_names[0]
    sheet_index = sheet_names.index(default_sheet) if default_sheet in sheet_names else 0
    sheet_name = st.selectbox("2. Escolha a aba", sheet_names, index=sheet_index)

    try:
        headers = detect_header_candidates(source_bytes, sheet_name)
        sheet_map = read_sheet_map(source_bytes, sheet_name, max_rows=40)
    except Exception as exc:
        st.error(f"Não foi possível analisar a aba: {exc}")
        return

    st.subheader("Mapa da planilha")
    st.dataframe(sheet_map, use_container_width=True, hide_index=True, height=260)

    if not headers:
        st.warning("Nenhum cabeçalho foi encontrado nas primeiras linhas. Informe uma linha de cabeçalho manualmente.")
        headers = detect_header_candidates(source_bytes, sheet_name, scan_rows=100, include_blank=True)
    if not headers:
        st.error("Esta aba não possui células que possam ser usadas como cabeçalho.")
        return

    labels = [header["label"] for header in headers]
    label_to_header = {header["label"]: header for header in headers}
    model_key = model_fingerprint(sheet_name, headers)
    preferred_reference = remembered.get("reference_label") if remembered else labels[0]
    preferred_destination = remembered.get("destination_label") if remembered else labels[min(1, len(labels) - 1)]
    ref_index = labels.index(preferred_reference) if preferred_reference in labels else 0
    dest_index = labels.index(preferred_destination) if preferred_destination in labels else min(1, len(labels) - 1)

    st.subheader("3. Configure o preenchimento")
    left, middle, right, fourth = st.columns(4)
    with left:
        reference_label = st.selectbox("Coluna de referência", labels, index=ref_index)
    with middle:
        destination_label = st.selectbox("Coluna de destino", labels, index=dest_index)
    with right:
        default_start = int(remembered.get("start_row", 0)) if remembered else 0
        start_row = st.number_input("Primeira linha de dados", min_value=1, value=default_start or label_to_header[reference_label]["row"] + 1)
    with fourth:
        unit_choices = [label for label in labels if label not in {reference_label, destination_label}]
        default_unit = remembered.get("unit_label") if remembered else (unit_choices[0] if unit_choices else reference_label)
        unit_label = st.selectbox("Coluna da unidade", unit_choices or labels,
                                  index=(unit_choices or labels).index(default_unit) if default_unit in (unit_choices or labels) else 0)

    use_date = st.checkbox("Preencher uma célula com a data", value=bool(remembered.get("date_cell")) if remembered else False)
    date_cell = ""
    selected_date = date.today()
    if use_date:
        date_col, date_value = st.columns(2)
        with date_col:
            date_cell = st.text_input("Célula da data", value=remembered.get("date_cell", "") if remembered else "", placeholder="Ex.: C1").strip().upper()
        with date_value:
            selected_date = st.date_input("Data", value=date.today())

    if st.button("Salvar esta configuração", use_container_width=False):
        save_template(source_key, {
            "sheet_name": sheet_name,
            "reference_label": reference_label,
            "destination_label": destination_label,
            "unit_label": unit_label,
            "start_row": int(start_row),
            "date_cell": date_cell,
        })
        st.success("Configuração salva para este modelo de planilha.")

    st.subheader("Cadastro de unidades")
    current_units = load_units(model_key)
    if not current_units:
        current_units = read_units(source_bytes, sheet_name, label_to_header[unit_label]["column"],
                                   label_to_header[reference_label]["column"], int(start_row))
        if current_units:
            save_units(model_key, current_units)
    with st.expander("Adicionar, editar ou desativar unidades", expanded=False):
        add_left, add_middle, add_right = st.columns([2, 1, 1])
        with add_left:
            new_unit = st.text_input("Nome da nova unidade", placeholder="Ex.: CO MACAÉ", key="new_unit")
        with add_middle:
            new_code = st.text_input("Código", placeholder="Ex.: S75", key="new_code").upper().strip()
        with add_right:
            st.write("")
            st.write("")
            add_clicked = st.button("Adicionar unidade")
        if add_clicked:
            if not new_unit.strip() or not new_code:
                st.error("Informe o nome da unidade e o código.")
            elif any(str(item.get("Código", "")).casefold() == new_code.casefold() for item in current_units):
                st.error("Este código já existe no cadastro.")
            else:
                current_units.append({"Unidade": new_unit.strip().upper(), "Código": new_code, "Ativa": True})
                save_units(model_key, current_units)
                st.success("Unidade cadastrada. Ela entrará na próxima cópia gerada.")
                st.rerun()
        catalog = pd.DataFrame(current_units, columns=["Unidade", "Código", "Ativa"])
        edited_catalog = st.data_editor(catalog, num_rows="dynamic", use_container_width=True, hide_index=True,
                                        column_config={"Ativa": st.column_config.CheckboxColumn("Ativa")}, key=f"catalog_{model_key}")
        if st.button("Salvar cadastro de unidades"):
            clean_units = []
            for item in edited_catalog.to_dict("records"):
                unit, code = str(item.get("Unidade", "")).strip(), str(item.get("Código", "")).strip().upper()
                if unit and code:
                    clean_units.append({"Unidade": unit, "Código": code, "Ativa": bool(item.get("Ativa", True))})
            save_units(model_key, clean_units)
            st.success("Cadastro atualizado. Desmarque 'Ativa' para excluir uma unidade das próximas cópias.")
            current_units = clean_units

    st.subheader("4. Cole os dados")
    raw_text = st.text_area(
        "Um registro por linha — exemplos: S21 - 6, S31: 8, Produto A = 15",
        height=150,
        placeholder="S21 - 6\nS31: 8\nS49 12",
    )
    if st.button("Interpretar dados", type="primary"):
        rows, warnings = parse_entries(raw_text)
        st.session_state.parsed_rows = rows
        st.session_state.parse_warnings = warnings
        if not rows:
            st.warning("Nenhum par referência/valor foi identificado. Revise o texto informado.")

    for warning in st.session_state.parse_warnings:
        st.warning(warning)

    if not st.session_state.parsed_rows:
        return

    st.subheader("5. Revise antes de gerar")
    preview = pd.DataFrame(st.session_state.parsed_rows, columns=["Referência", "Valor"])
    edited = st.data_editor(preview, num_rows="dynamic", use_container_width=True, hide_index=True, key="preview_editor")

    matching_references = find_matching_references(
        source_bytes, sheet_name, label_to_header[reference_label]["column"], int(start_row), edited.to_dict("records")
    )
    if matching_references:
        st.success(f"Configuração validada: {len(matching_references)} referência(s) será(ão) encontrada(s) na coluna selecionada.")
    else:
        st.error(
            "Nenhuma referência da prévia foi encontrada na coluna selecionada. "
            "Para esta planilha, escolha B — CÓDIGO como referência, C — QUANTIDADE DE PALETES como destino e linha 3."
        )

    if st.button("Gerar arquivos", type="primary"):
        if not matching_references:
            st.error("A geração foi cancelada para não baixar uma planilha vazia. Corrija a configuração acima.")
            return
        try:
            result = build_filled_workbook(
                source_bytes=source_bytes,
                sheet_name=sheet_name,
                reference_column=label_to_header[reference_label]["column"],
                destination_column=label_to_header[destination_label]["column"],
                start_row=int(start_row),
                entries=edited.to_dict("records"),
                date_cell=date_cell or None,
                date_value=selected_date if date_cell else None,
                unit_column=label_to_header[unit_label]["column"],
                units=load_units(model_key),
            )
        except ValueError as exc:
            st.error(str(exc))
            return
        except Exception as exc:
            st.error(f"Não foi possível gerar a cópia: {exc}")
            return

        try:
            pdf_bytes = convert_workbook_to_pdf(result.workbook_bytes, sheet_name)
        except PdfConversionError as exc:
            st.warning(f"PDF fiel ao Excel indisponível: {exc}")
            pdf_bytes = generate_pdf(
                title="SmartSheet — relatório de preenchimento",
                sheet_name=sheet_name,
                rows=edited.to_dict("records"),
                generated_date=selected_date,
            )
        st.session_state.generated_output = {
            "excel": result.workbook_bytes,
            "pdf": pdf_bytes,
            "name_base": safe_output_name(Path(uploaded.name).stem),
            "updated_count": result.updated_count,
            "missing": result.missing_references,
        }

    generated = st.session_state.generated_output
    if not generated:
        return
    st.success(f"Cópia pronta: {generated['updated_count']} célula(s) atualizada(s).")
    if generated["missing"]:
        st.warning("Referências não encontradas: " + ", ".join(generated["missing"]))
    st.download_button("⬇️ Baixar Excel preenchido", generated["excel"], f"{generated['name_base']}_preenchida.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.download_button("⬇️ Baixar PDF da planilha", generated["pdf"], f"{generated['name_base']}_preenchida.pdf",
                       "application/pdf", use_container_width=True)
    try:
        st.subheader("Visualizações")
        preview_tab, pdf_tab, print_tab = st.tabs(["Planilha", "PDF", "Tirar print"])
        with preview_tab:
            st.caption("Esta é a primeira página da planilha preenchida, renderizada pelo Excel.")
            zoom_out, zoom_label, zoom_in = st.columns([1, 2, 1])
            with zoom_out:
                if st.button("− Afastar", key="zoom_out"):
                    st.session_state.preview_zoom = max(0.6, st.session_state.preview_zoom - 0.2)
                    st.rerun()
            with zoom_label:
                st.markdown(f"<div style='text-align:center; padding-top:8px'>Zoom: {int(st.session_state.preview_zoom * 100)}%</div>", unsafe_allow_html=True)
            with zoom_in:
                if st.button("+ Aproximar", key="zoom_in"):
                    st.session_state.preview_zoom = min(2.0, st.session_state.preview_zoom + 0.2)
                    st.rerun()
            preview_image = render_pdf_page(generated["pdf"], crop_to_content=True, zoom=st.session_state.preview_zoom)
            st.image(preview_image, width=int(700 * st.session_state.preview_zoom))
        with pdf_tab:
            encoded_pdf = base64.b64encode(generated["pdf"]).decode("utf-8")
            components.html(f'<iframe src="data:application/pdf;base64,{encoded_pdf}" width="100%" height="720" type="application/pdf"></iframe>', height=740)
        with print_tab:
            page_count = pdf_page_count(generated["pdf"])
            page_number = st.number_input("Página para o print", min_value=1, max_value=page_count, value=1, step=1)
            screenshot = render_pdf_page(generated["pdf"], int(page_number) - 1, crop_to_content=True)
            st.image(screenshot, use_container_width=True)
            st.download_button("⬇️ Baixar print em PNG", screenshot, f"{generated['name_base']}_pagina_{page_number}.png", "image/png")
    except ImportError:
        st.warning("Para habilitar a visualização e o print, instale a dependência PyMuPDF e reinicie o SmartSheet.")


if __name__ == "__main__":
    main()
