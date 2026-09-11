import re
import unicodedata
from typing import Any

import pandas as pd

from config import PAINT_EXCEL_FILE


PAINT_FILTER_COLUMNS = [
    "Ano",
    "Auditoria",
    "Processo",
    "Realização",
    "Status",
]

PAINT_STATUS_ORDER = ["Concluída", "Em Andamento", "Não Iniciada"]
def _normalize_text(value: Any) -> str:
    text = "" if pd.isna(value) else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", text).strip().lower()


def _clean_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _detect_header_row(file_path, sheet_name: str) -> int:
    preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=20)
    required = {_normalize_text(value) for value in PAINT_FILTER_COLUMNS}

    for row_index, row in preview.iterrows():
        values = {_normalize_text(value) for value in row if not pd.isna(value)}
        if required.issubset(values):
            return int(row_index)

    raise ValueError(
        "Não foi possível localizar o cabeçalho da planilha de Execução do PAINT."
    )


def _normalize_status(value: Any) -> str:
    normalized = _normalize_text(value)
    labels = {
        "concluida": "Concluída",
        "concluido": "Concluída",
        "em andamento": "Em Andamento",
        "nao iniciada": "Não Iniciada",
        "nao iniciado": "Não Iniciada",
    }
    return labels.get(normalized, _clean_header(value))


def load_paint_data() -> pd.DataFrame:
    if not PAINT_EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"A base de Execução do PAINT não foi encontrada em: {PAINT_EXCEL_FILE}"
        )

    workbook = pd.ExcelFile(PAINT_EXCEL_FILE)
    sheet_name = "Execução" if "Execução" in workbook.sheet_names else workbook.sheet_names[0]
    header_row = _detect_header_row(PAINT_EXCEL_FILE, sheet_name)
    df = pd.read_excel(PAINT_EXCEL_FILE, sheet_name=sheet_name, header=header_row)
    df.columns = [_clean_header(column) for column in df.columns]
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")

    missing = [column for column in ["Ano", "N.", *PAINT_FILTER_COLUMNS[1:]] if column not in df.columns]
    if missing:
        raise ValueError(
            "A planilha de Execução do PAINT não possui os campos necessários: "
            + ", ".join(missing)
        )

    df = df[df["Ano"].notna()].copy()
    df["Ano"] = pd.to_numeric(df["Ano"], errors="coerce").astype("Int64")
    df["N."] = pd.to_numeric(df["N."], errors="coerce").astype("Int64")
    df["Status"] = df["Status"].map(_normalize_status)

    for column in ["Auditoria", "Processo", "Realização"]:
        df[column] = df[column].apply(_clean_header)

    df.reset_index(drop=True, inplace=True)
    return df[["Ano", "N.", "Auditoria", "Processo", "Realização", "Status"]]


def apply_paint_filters(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    filtered = df.copy()
    for column in PAINT_FILTER_COLUMNS:
        value = filters.get(column)
        if not value:
            continue
        selected = filtered[column].map(_normalize_text) == _normalize_text(value)
        filtered = filtered.loc[selected]
    return filtered.reset_index(drop=True)


def get_paint_filter_options(df: pd.DataFrame) -> dict:
    options = {}
    for column in PAINT_FILTER_COLUMNS:
        values = {
            _clean_header(value)
            for value in df[column].dropna()
            if _clean_header(value)
        }
        if column == "Ano":
            options[column] = sorted(values, key=lambda value: int(float(value)))
        elif column == "Status":
            options[column] = [status for status in PAINT_STATUS_ORDER if status in values]
        else:
            options[column] = sorted(values, key=lambda value: _normalize_text(value))
    return options


def get_paint_indicators(df: pd.DataFrame) -> dict:
    status = df["Status"].map(_normalize_status)
    counts = {
        label: int((status == label).sum())
        for label in PAINT_STATUS_ORDER
    }
    total = int(len(df))

    return {
        "total_auditorias": total,
        "concluidas": counts["Concluída"],
        "em_andamento": counts["Em Andamento"],
        "nao_iniciadas": counts["Não Iniciada"],
        "distribuicao_status": {
            label: {
                "quantidade": counts[label],
                "percentual": round(counts[label] / total * 100, 2) if total else 0,
            }
            for label in PAINT_STATUS_ORDER
        },
        "processos_status": [
            {
                "numero": int(row["N."]) if not pd.isna(row["N."]) else index + 1,
                "processo": _clean_header(row["Processo"]),
                "status": _normalize_status(row["Status"]),
            }
            for index, row in df.iterrows()
        ],
    }


def get_paint_source_summary(df: pd.DataFrame) -> dict:
    years = sorted({int(value) for value in df["Ano"].dropna()})
    if not years:
        period = "Não informado"
    elif len(years) == 1:
        period = str(years[0])
    else:
        period = f"{years[0]} a {years[-1]}"

    preview = pd.read_excel(PAINT_EXCEL_FILE, sheet_name=0, header=None, nrows=3)
    position = ""
    for value in preview.to_numpy().flatten():
        match = re.search(r"Posição\s*:\s*(.+)", _clean_header(value), flags=re.IGNORECASE)
        if match:
            position = match.group(1).strip()
            break

    return {
        "arquivo": PAINT_EXCEL_FILE.name,
        "periodo": period,
        "posicao": position,
    }


def paint_dataframe_records(df: pd.DataFrame) -> list[dict]:
    records = []
    for row in df.to_dict(orient="records"):
        record = {}
        for column, value in row.items():
            if pd.isna(value):
                record[column] = ""
            elif hasattr(value, "item"):
                record[column] = value.item()
            else:
                record[column] = value
        records.append(record)
    return records
