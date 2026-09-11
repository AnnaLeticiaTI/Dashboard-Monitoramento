import re
import unicodedata
from datetime import date, timedelta
from typing import Any

import pandas as pd

from config import CARTEIRA_EXCEL_FILE


CARTEIRA_FILTER_COLUMNS = ["Auditoria", "PAINT", "Status"]
CARTEIRA_STATUS_ORDER = [
    "Planos de ação em monitoramento",
    "Encerrada",
    "Aguardando plano de ação",
]
def _normalize_text(value: Any) -> str:
    text = "" if pd.isna(value) else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", text).strip().lower()


def _clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _detect_header_row(sheet_name: str, required_columns: list[str]) -> int:
    preview = pd.read_excel(
        CARTEIRA_EXCEL_FILE,
        sheet_name=sheet_name,
        header=None,
        nrows=45,
    )
    required = {_normalize_text(column) for column in required_columns}

    for row_index, row in preview.iterrows():
        values = {_normalize_text(value) for value in row if not pd.isna(value)}
        if required.issubset(values):
            return int(row_index)

    raise ValueError(
        f"Não foi possível localizar o cabeçalho da aba {sheet_name}."
    )


def _normalize_carteira_status(value: Any) -> str:
    normalized = _normalize_text(value)
    if "monitoramento" in normalized:
        return "Planos de ação em monitoramento"
    if "encerrad" in normalized:
        return "Encerrada"
    if "aguardando" in normalized:
        return "Aguardando plano de ação"
    return _clean_text(value)


def _easter_sunday(year: int) -> date:
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def _brazilian_holidays(start_year: int, end_year: int) -> list[date]:
    holidays = []
    for year in range(start_year, end_year + 1):
        easter = _easter_sunday(year)
        holidays.extend(
            [
                date(year, 1, 1),
                easter - timedelta(days=48),
                easter - timedelta(days=47),
                easter - timedelta(days=2),
                date(year, 4, 21),
                date(year, 5, 1),
                easter + timedelta(days=60),
                date(year, 9, 7),
                date(year, 10, 12),
                date(year, 11, 2),
                date(year, 11, 15),
                date(year, 12, 25),
            ]
        )
        if year >= 2024:
            holidays.append(date(year, 11, 20))
    return holidays


def _business_deadline(value: Any) -> pd.Timestamp:
    request_date = pd.to_datetime(value, errors="coerce")
    if pd.isna(request_date):
        return pd.NaT
    holidays = _brazilian_holidays(request_date.year, request_date.year + 1)
    business_days = pd.offsets.CustomBusinessDay(n=15, holidays=holidays)
    return (request_date.normalize() + business_days).normalize()


def _calculate_aguardando_deadlines(dataframe: pd.DataFrame) -> pd.DataFrame:
    dataframe = dataframe.copy()
    request_column = "Data da solicitação plano ação"
    deadline_column = "Data limite para entrega do plano de ação*"
    delivery_column = "Data da entrega do plano de ação"

    request_dates = dataframe[request_column]
    delivery_dates = dataframe[delivery_column]
    deadline_dates = request_dates.apply(_business_deadline)
    has_request = request_dates.notna()
    has_delivery = delivery_dates.notna()
    overdue = has_request & ~has_delivery & (pd.Timestamp.today().normalize() > deadline_dates)

    statuses = dataframe["Status"].apply(_clean_text)
    normalized_statuses = statuses.map(_normalize_text)
    statuses.loc[has_request & ~has_delivery & ~overdue] = "No Prazo"
    statuses.loc[overdue] = "Vencido"
    statuses.loc[
        has_delivery
        & normalized_statuses.isin({"", "vencido"})
    ] = "Entregue"

    dataframe[deadline_column] = deadline_dates
    dataframe["Status"] = statuses
    dataframe["Dias"] = (
        delivery_dates.sub(request_dates).dt.days.where(has_delivery).astype("Int64")
    )
    return dataframe


def _clean_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    dataframe = dataframe.copy()
    dataframe.columns = [_clean_text(column) for column in dataframe.columns]
    unnamed = [
        column
        for column in dataframe.columns
        if _normalize_text(column).startswith("unnamed")
    ]
    return (
        dataframe
        .drop(columns=unnamed, errors="ignore")
        .dropna(axis=1, how="all")
        .dropna(axis=0, how="all")
    )


def load_carteira_data() -> pd.DataFrame:
    if not CARTEIRA_EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"A base da Carteira de Auditoria não foi encontrada em: {CARTEIRA_EXCEL_FILE}"
        )

    workbook = pd.ExcelFile(CARTEIRA_EXCEL_FILE)
    sheet_name = "Carteira" if "Carteira" in workbook.sheet_names else workbook.sheet_names[0]
    header_row = _detect_header_row(sheet_name, CARTEIRA_FILTER_COLUMNS)
    dataframe = pd.read_excel(
        CARTEIRA_EXCEL_FILE,
        sheet_name=sheet_name,
        header=header_row,
    )
    dataframe = _clean_columns(dataframe)

    required = [
        "Auditoria",
        "PAINT",
        "Status",
        "Total Planos de Ação",
        "% Implementação",
        "Planos Abertos",
    ]
    missing = [column for column in required if column not in dataframe.columns]
    if missing:
        raise ValueError(
            "A aba Carteira não possui os campos necessários: " + ", ".join(missing)
        )

    dataframe = dataframe[dataframe["Auditoria"].notna()].copy()
    dataframe["Auditoria"] = dataframe["Auditoria"].apply(_clean_text)
    dataframe["PAINT"] = pd.to_numeric(dataframe["PAINT"], errors="coerce").astype("Int64")
    dataframe["Status"] = dataframe["Status"].map(_normalize_carteira_status)
    for column in ["Total Planos de Ação", "% Implementação", "Planos Abertos"]:
        dataframe[column] = pd.to_numeric(dataframe[column], errors="coerce")

    dataframe.reset_index(drop=True, inplace=True)
    return dataframe[required]


def load_aguardando_pa_data() -> pd.DataFrame:
    workbook = pd.ExcelFile(CARTEIRA_EXCEL_FILE)
    if "Aguard. PA" not in workbook.sheet_names:
        raise ValueError("A planilha da Carteira não possui a aba Aguard. PA.")

    header_row = _detect_header_row(
        "Aguard. PA",
        ["Auditorias", "PAINT", "Status"],
    )
    dataframe = pd.read_excel(
        CARTEIRA_EXCEL_FILE,
        sheet_name="Aguard. PA",
        header=header_row,
    )
    dataframe = _clean_columns(dataframe)

    required = [
        "Auditorias",
        "PAINT",
        "Data da solicitação plano ação",
        "Data limite para entrega do plano de ação*",
        "Status",
        "Data da entrega do plano de ação",
        "Dias",
    ]
    missing = [column for column in required if column not in dataframe.columns]
    if missing:
        raise ValueError(
            "A aba Aguard. PA não possui os campos necessários: " + ", ".join(missing)
        )

    dataframe = dataframe[
        dataframe["Auditorias"].notna() & dataframe["PAINT"].notna()
    ].copy()
    dataframe["Auditorias"] = dataframe["Auditorias"].apply(_clean_text)
    dataframe["PAINT"] = pd.to_numeric(dataframe["PAINT"], errors="coerce").astype("Int64")
    for column in [
        "Data da solicitação plano ação",
        "Data limite para entrega do plano de ação*",
        "Data da entrega do plano de ação",
    ]:
        dataframe[column] = pd.to_datetime(dataframe[column], errors="coerce")
    dataframe = _calculate_aguardando_deadlines(dataframe)

    dataframe.reset_index(drop=True, inplace=True)
    return dataframe[required]


def apply_carteira_filters(dataframe: pd.DataFrame, filters: dict) -> pd.DataFrame:
    filtered = dataframe.copy()
    for column in CARTEIRA_FILTER_COLUMNS:
        value = filters.get(column)
        if not value:
            continue
        selected = filtered[column].map(_normalize_text) == _normalize_text(value)
        filtered = filtered.loc[selected]
    return filtered.reset_index(drop=True)


def apply_aguardando_filters(
    dataframe: pd.DataFrame,
    filtered_carteira: pd.DataFrame,
) -> pd.DataFrame:
    allowed_pairs = {
        (_normalize_text(row["Auditoria"]), _normalize_text(row["PAINT"]))
        for _, row in filtered_carteira.iterrows()
    }
    if not allowed_pairs:
        return dataframe.iloc[0:0].copy()

    selected = dataframe.apply(
        lambda row: (
            _normalize_text(row["Auditorias"]),
            _normalize_text(row["PAINT"]),
        ) in allowed_pairs,
        axis=1,
    )
    return dataframe.loc[selected].reset_index(drop=True)


def get_carteira_filter_options(dataframe: pd.DataFrame) -> dict:
    options = {}
    for column in CARTEIRA_FILTER_COLUMNS:
        values = {
            _clean_text(value)
            for value in dataframe[column].dropna()
            if _clean_text(value)
        }
        if column == "PAINT":
            options[column] = sorted(values, key=lambda value: int(float(value)))
        elif column == "Status":
            options[column] = [status for status in CARTEIRA_STATUS_ORDER if status in values]
        else:
            options[column] = sorted(values, key=lambda value: _normalize_text(value))
    return options


def get_carteira_indicators(dataframe: pd.DataFrame) -> dict:
    statuses = dataframe["Status"].map(_normalize_carteira_status)
    total = int(len(dataframe))
    monitoring = int((statuses == "Planos de ação em monitoramento").sum())
    closed = int((statuses == "Encerrada").sum())
    awaiting = int((statuses == "Aguardando plano de ação").sum())

    status_counts = {
        "Auditorias com Planos em Monitoramento": monitoring,
        "Auditorias Encerradas": closed,
        "Aguardando Plano de Ação": awaiting,
    }
    year_counts = (
        dataframe["PAINT"]
        .dropna()
        .astype(int)
        .value_counts()
        .sort_index()
    )

    return {
        "total_auditorias": total,
        "planos_monitoramento": monitoring,
        "encerradas": closed,
        "aguardando_plano": awaiting,
        "distribuicao_status": {
            label: {
                "quantidade": quantity,
                "percentual": round(quantity / total * 100, 2) if total else 0,
            }
            for label, quantity in status_counts.items()
        },
        "distribuicao_paint": {
            str(year): {
                "quantidade": int(quantity),
                "percentual": round(int(quantity) / total * 100, 2) if total else 0,
            }
            for year, quantity in year_counts.items()
        },
    }


def get_carteira_source_summary(dataframe: pd.DataFrame) -> dict:
    years = sorted({int(value) for value in dataframe["PAINT"].dropna()})
    if not years:
        period = "Não informado"
    elif len(years) == 1:
        period = str(years[0])
    else:
        period = f"{years[0]} a {years[-1]}"
    return {
        "arquivo": CARTEIRA_EXCEL_FILE.name,
        "periodo": period,
    }


def carteira_dataframe_records(dataframe: pd.DataFrame) -> list[dict]:
    records = []
    for row in dataframe.to_dict(orient="records"):
        record = {}
        for column, value in row.items():
            if pd.isna(value):
                record[column] = ""
            elif isinstance(value, pd.Timestamp):
                record[column] = value.strftime("%d/%m/%Y")
            elif column == "% Implementação":
                record[column] = f"{float(value):.1f}%".replace(".", ",")
            elif isinstance(value, float) and value.is_integer():
                record[column] = int(value)
            elif hasattr(value, "item"):
                record[column] = value.item()
            else:
                record[column] = value
        records.append(record)
    return records
