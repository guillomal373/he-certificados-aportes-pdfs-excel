#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


SHEET_COLUMNS = {
    "Liquidaciones Pagadas": [
        "archivo_pdf",
        "nombre_persona",
        "tipo_id",
        "numero_id",
        "periodo_pension",
        "periodo_salud",
        "tipo_planilla",
        "clave",
        "no_transaccion",
        "fecha_pago",
        "fecha_generacion_certificado",
    ],
    "Seguridad Social": [
        "archivo_pdf",
        "nombre_persona",
        "tipo_id",
        "numero_id",
        "periodo_pension",
        "tipo_planilla",
        "afp_administradora",
        "afp_dias",
        "afp_ibc",
        "afp_fs",
        "afp_fsp",
        "afp_aporte",
        "eps_administradora",
        "eps_dias",
        "eps_ibc",
        "eps_aporte",
        "arl_administradora",
        "arl_dias",
        "arl_ibc",
        "arl_aporte",
        "fecha_generacion_certificado",
    ],
    "Aportes Parafiscales": [
        "archivo_pdf",
        "nombre_persona",
        "tipo_id",
        "numero_id",
        "periodo_pension",
        "tipo_planilla",
        "ccf_administradora",
        "ccf_dias",
        "ccf_ibc",
        "ccf_aporte",
        "otros_parafiscales_ibc",
        "icbf_tarifa",
        "icbf_aporte",
        "sena_tarifa",
        "sena_aporte",
        "esap_tarifa",
        "esap_aporte",
        "men_tarifa",
        "men_aporte",
        "fecha_generacion_certificado",
    ],
    "Novedades": [
        "archivo_pdf",
        "nombre_persona",
        "tipo_id",
        "numero_id",
        "periodo_pension",
        "tipo_planilla",
        "codigo_novedad",
        "marcado",
        "fecha_generacion_certificado",
    ],
}

NOVEDADES_CODIGOS = [
    "ING", "RET", "TAE", "TDE", "TAP",
    "TDP", "VSP", "VST", "SLN", "IGE",
    "LMA", "VAC", "AVP", "IRL", "VCT",
]

SECURITY_PATTERN = re.compile(
    r"(\d{4}/\d{2})\s+([A-Z])\s+([A-Z ]+?)\s+(\d+)\s+\$([\d,]+)\s+\$([\d,]+)\s+\$([\d,]+)\s+\$([\d,]+)\s+"
    r"([A-Z ]+?)\s+(\d+)\s+\$([\d,]+)\s+\$([\d,]+)\s+([A-Z ]+?)\s+(\d+)\s+\$([\d,]+)\s+\$([\d,]+)"
)

PARAF_PATTERN = re.compile(
    r"(\d{4}/\d{2})\s+([A-Z])\s+([A-Z ]+?)\s+(\d+)\s+\$([\d,]+)\s+\$([\d,]+)\s+\$([\d,]+)\s+"
    r"(\d+)%\s+\$([\d,]+)\s+(\d+)%\s+\$([\d,]+)\s+(\d+)%\s+\$([\d,]+)\s+(\d+)%\s+\$([\d,]+)"
)

LIQ_PATTERN = re.compile(
    r"(\d{4}/\d{2})\s+(\d{4}/\d{2})\s+([A-Z])\s+(\d+)\s+(\d+)\s+(\d{4}/\d{2}/\d{2})"
)

HEADER_PATTERN = re.compile(
    r"Se certifica que\s+.*?para\s+(.+?)\s+identificado con\s+([A-Z]{2})\s+(\d+)\s*:?",
    re.IGNORECASE | re.DOTALL,
)

GEN_DATE_PATTERN = re.compile(
    r"Certificado generado el\s+(\d{4}-\d{2}-\d{2})\s+a las\s+(\d{2}:\d{2})",
    re.IGNORECASE,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Procesa certificados de aportes PDF en una carpeta y genera un Excel consolidado."
    )
    parser.add_argument("--input", required=True, help="Carpeta que contiene los PDF.")
    parser.add_argument("--output", required=True, help="Ruta completa del XLSX de salida.")
    return parser.parse_args()


def clean_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def money_to_int(value: str) -> int:
    return int(re.sub(r"[^\d]", "", value))


def parse_date(value: str) -> datetime:
    return datetime.strptime(value.replace("-", "/"), "%Y/%m/%d")


def normalize_admin_name(value: str) -> str:
    raw = clean_spaces(value.upper())
    replacements = {
        "ARLSURA": "ARL SURA",
        "COLPATRIAARP": "COLPATRIA ARP",
        "SALUDTOTAL": "SALUD TOTAL",
    }
    return replacements.get(raw, raw)


def crop_text(page) -> str:
    candidates = []

    # 1) Texto completo sin recorte
    try:
        text_full = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
        candidates.append(text_full)
    except Exception:
        pass

    # 2) Recorte desde más abajo, como venías haciendo
    try:
        cropped1 = page.crop((0, 100, page.width, page.height - 20))
        text_c1 = cropped1.extract_text(x_tolerance=2, y_tolerance=3) or ""
        candidates.append(text_c1)
    except Exception:
        pass

    # 3) Recorte un poco más conservador
    try:
        cropped2 = page.crop((0, 60, page.width, page.height - 20))
        text_c2 = cropped2.extract_text(x_tolerance=2, y_tolerance=3) or ""
        candidates.append(text_c2)
    except Exception:
        pass

    # 4) Intento con palabras individuales reconstruidas
    try:
        words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
        text_words = " ".join(w["text"] for w in words)
        candidates.append(text_words)
    except Exception:
        pass

    # Normalizar espacios
    normalized = [clean_spaces(t) for t in candidates if t and t.strip()]

    # Priorizar el texto que sí contenga el encabezado real
    for text in normalized:
        if "Se certifica que" in text:
            return text[text.find("Se certifica que"):]

    # Si ninguno trae el encabezado, devolver el más largo para debug
    if normalized:
        return max(normalized, key=len)

    return ""


def extract_last_identity(full_text: str) -> Tuple[str, str, int]:
    matches = HEADER_PATTERN.findall(full_text)
    if not matches:
        raise ValueError(f"No se pudo extraer la identidad del encabezado. Texto detectado: {full_text[:500]}")
    person_name, id_type, id_number = matches[-1]
    return clean_spaces(person_name), id_type.upper(), int(id_number)


def extract_generation_date(full_text: str) -> datetime:
    match = GEN_DATE_PATTERN.search(full_text)
    if not match:
        raise ValueError("No se pudo extraer la fecha de generación.")
    date_part, time_part = match.groups()
    return datetime.strptime(f"{date_part} {time_part}", "%Y-%m-%d %H:%M")


def get_section(full_text: str, start: str, end: str | None) -> str:
    start_idx = full_text.find(start)
    if start_idx == -1:
        return ""
    if end is None:
        return full_text[start_idx:]
    end_idx = full_text.find(end, start_idx)
    if end_idx == -1:
        return full_text[start_idx:]
    return full_text[start_idx:end_idx]


def extract_tables(page):
    tables = page.extract_tables()
    if len(tables) < 4:
        raise ValueError(f"Se esperaban al menos 4 tablas y se encontraron {len(tables)}.")
    return tables


def parse_liquidaciones(
    section_text: str,
    archivo_pdf: str,
    nombre: str,
    tipo_id: str,
    numero_id: int,
    fecha_generacion: datetime,
) -> List[List]:
    rows = []
    for periodo_pension, periodo_salud, tipo_planilla, clave, no_transaccion, fecha_pago in LIQ_PATTERN.findall(section_text):
        rows.append([
            archivo_pdf,
            nombre,
            tipo_id,
            numero_id,
            periodo_pension,
            periodo_salud,
            tipo_planilla,
            int(clave),
            int(no_transaccion),
            parse_date(fecha_pago),
            fecha_generacion,
        ])
    return rows


def parse_seguridad_social(
    section_text: str,
    archivo_pdf: str,
    nombre: str,
    tipo_id: str,
    numero_id: int,
    fecha_generacion: datetime,
) -> List[List]:
    rows = []
    for match in SECURITY_PATTERN.findall(section_text):
        (
            periodo_pension, tipo_planilla,
            afp_adm, afp_dias, afp_ibc, afp_fs, afp_fsp, afp_aporte,
            eps_adm, eps_dias, eps_ibc, eps_aporte,
            arl_adm, arl_dias, arl_ibc, arl_aporte,
        ) = match

        rows.append([
            archivo_pdf,
            nombre,
            tipo_id,
            numero_id,
            periodo_pension,
            tipo_planilla,
            normalize_admin_name(afp_adm),
            int(afp_dias),
            money_to_int(afp_ibc),
            money_to_int(afp_fs),
            money_to_int(afp_fsp),
            money_to_int(afp_aporte),
            normalize_admin_name(eps_adm),
            int(eps_dias),
            money_to_int(eps_ibc),
            money_to_int(eps_aporte),
            normalize_admin_name(arl_adm),
            int(arl_dias),
            money_to_int(arl_ibc),
            money_to_int(arl_aporte),
            fecha_generacion,
        ])
    return rows


def parse_parafiscales(
    section_text: str,
    archivo_pdf: str,
    nombre: str,
    tipo_id: str,
    numero_id: int,
    fecha_generacion: datetime,
) -> List[List]:
    rows = []
    for match in PARAF_PATTERN.findall(section_text):
        (
            periodo_pension, tipo_planilla,
            ccf_adm, ccf_dias, ccf_ibc, ccf_aporte, otros_ibc,
            icbf_tarifa, icbf_aporte,
            sena_tarifa, sena_aporte,
            esap_tarifa, esap_aporte,
            men_tarifa, men_aporte,
        ) = match

        rows.append([
            archivo_pdf,
            nombre,
            tipo_id,
            numero_id,
            periodo_pension,
            tipo_planilla,
            normalize_admin_name(ccf_adm),
            int(ccf_dias),
            money_to_int(ccf_ibc),
            money_to_int(ccf_aporte),
            money_to_int(otros_ibc),
            int(icbf_tarifa),
            money_to_int(icbf_aporte),
            int(sena_tarifa),
            money_to_int(sena_aporte),
            int(esap_tarifa),
            money_to_int(esap_aporte),
            int(men_tarifa),
            money_to_int(men_aporte),
            fecha_generacion,
        ])
    return rows


def parse_novedades_from_table(
    table,
    archivo_pdf: str,
    nombre: str,
    tipo_id: str,
    numero_id: int,
    fecha_generacion: datetime,
) -> List[List]:
    rows = []
    data_rows = table[2:]

    for row in data_rows:
        if not row or len(row) < 17:
            continue

        periodo_match = re.search(r"\d{4}/\d{2}", row[0] or "")
        tipo_match = re.search(r"\b[A-Z]\b", clean_spaces(row[1] or ""))

        if not periodo_match or not tipo_match:
            continue

        periodo_pension = periodo_match.group(0)
        tipo_planilla = tipo_match.group(0)

        for idx, codigo in enumerate(NOVEDADES_CODIGOS, start=2):
            cell_value = clean_spaces((row[idx] or "").upper())
            if "X" in cell_value:
                rows.append([
                    archivo_pdf,
                    nombre,
                    tipo_id,
                    numero_id,
                    periodo_pension,
                    tipo_planilla,
                    codigo,
                    "X",
                    fecha_generacion,
                ])

    return rows


def process_pdf(pdf_path: Path) -> Dict[str, List[List]]:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        full_text = crop_text(page)
        print(f"DEBUG TEXTO {pdf_path.name}:")
        print(full_text[:800])
        print("-" * 80)
        tables = extract_tables(page)

    nombre, tipo_id, numero_id = extract_last_identity(full_text)
    fecha_generacion = extract_generation_date(full_text)

    section_liq = get_section(full_text, "Datos de las Liquidaciones Pagadas", "Aportes Sistema de Seguridad Social")
    section_seg = get_section(full_text, "Aportes Sistema de Seguridad Social", "Aportes Parafiscales")
    section_par = get_section(full_text, "Aportes Parafiscales", "Novedades")

    return {
        "Liquidaciones Pagadas": parse_liquidaciones(
            section_liq, pdf_path.name, nombre, tipo_id, numero_id, fecha_generacion
        ),
        "Seguridad Social": parse_seguridad_social(
            section_seg, pdf_path.name, nombre, tipo_id, numero_id, fecha_generacion
        ),
        "Aportes Parafiscales": parse_parafiscales(
            section_par, pdf_path.name, nombre, tipo_id, numero_id, fecha_generacion
        ),
        "Novedades": parse_novedades_from_table(
            tables[3], pdf_path.name, nombre, tipo_id, numero_id, fecha_generacion
        ),
    }


def build_workbook() -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for sheet_name, columns in SHEET_COLUMNS.items():
        ws = wb.create_sheet(sheet_name)
        for col_idx, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=1, column=col_idx, value=col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(len(col_name) + 3, 14), 28)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}1"
        ws.row_dimensions[1].height = 24

    return wb


def apply_formats(ws, sheet_name: str) -> None:
    if ws.max_row <= 1:
        return

    date_columns_by_sheet = {
        "Liquidaciones Pagadas": {10, 11},
        "Seguridad Social": {21},
        "Aportes Parafiscales": {20},
        "Novedades": {9},
    }

    numeric_columns_by_sheet = {
        "Liquidaciones Pagadas": {4, 8, 9},
        "Seguridad Social": {4, 8, 9, 10, 11, 12, 14, 15, 16, 18, 19, 20},
        "Aportes Parafiscales": {4, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19},
        "Novedades": {4},
    }

    date_cols = date_columns_by_sheet[sheet_name]
    numeric_cols = numeric_columns_by_sheet[sheet_name]

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            if cell.column in date_cols and cell.value is not None:
                if isinstance(cell.value, datetime) and cell.column in {11, 20, 21, 9}:
                    cell.number_format = "yyyy-mm-dd hh:mm"
                else:
                    cell.number_format = "yyyy-mm-dd"
            elif cell.column in numeric_cols and cell.value is not None:
                cell.number_format = "0"


def main() -> None:
    args = parse_args()
    input_dir = Path(args.input)
    output_file = Path(args.output)

    if not input_dir.exists() or not input_dir.is_dir():
        raise SystemExit(f"La carpeta de entrada no existe o no es válida: {input_dir}")

    pdf_files = sorted(input_dir.glob("*.pdf"))
    if not pdf_files:
        raise SystemExit(f"No se encontraron archivos PDF en: {input_dir}")

    consolidated = {sheet_name: [] for sheet_name in SHEET_COLUMNS}

    print(f"Procesando {len(pdf_files)} archivos PDF...")
    for pdf_path in pdf_files:
        try:
            data = process_pdf(pdf_path)
            for sheet_name, rows in data.items():
                consolidated[sheet_name].extend(rows)
            print(f"OK  - {pdf_path.name}")
        except Exception as exc:
            print(f"ERROR - {pdf_path.name}: {exc}")

    wb = build_workbook()

    for sheet_name, rows in consolidated.items():
        ws = wb[sheet_name]
        for row in rows:
            ws.append(row)
        apply_formats(ws, sheet_name)

    output_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_file)
    print(f"\\nExcel generado en: {output_file}")


if __name__ == "__main__":
    main()