import math
import re
import unicodedata
from typing import Any

import pandas as pd

from config import EXCEL_FILE

FILTER_COLUMNS = [
    "Origem",
    "Relatório de Auditoria",
    "Diretoria",
    "Unidade/Agência",
    "Responsável",
    "Classificação de Risco",
    "Status",
    "Ano",
]

SEMAFORO_FILTER_COLUMNS = [
    "Situação",
    "Classificação de Risco",
    "Status",
]

FILTER_COLUMN_MAP = {
    "Origem": ["Origem"],
    "Relatório de Auditoria": ["Relatório de Auditoria"],
    "Diretoria": ["Diretoria P.A.", "Diretoria"],
    "Unidade/Agência": [
        "Unidade Responsável P.A.",
        "Unidade/Agência",
        "Unidade Auditada",
        "Unidade",
    ],
    "Responsável": [
        "Responsável pelo plano de Ação",
        "Responsável pelo Plano de Ação",
        "Responsável",
    ],
    "Classificação de Risco": ["Classificação Risco", "Classificação de Risco"],
    "Status": ["Status"],
    "Ano": ["Ano"],
}

REINCIDENTE_COLUMN_ALIASES = [
    "Reincidente",
    "Reincidência",
    "Reincidencia",
]

ACTIVE_STATUS_KEYS = {"a vencer", "repactuado", "vencido"}

SEMAFORO_RULES = [
    ("vencidos_qualquer_risco", "Vencidos (qualquer risco)", "Atenção"),
    ("repactuados_risco_alto", "Repactuados + risco alto", "Atenção"),
    (
        "repactuados_risco_significativo",
        "Repactuados + risco significativo",
        "Atenção",
    ),
    (
        "repactuados_risco_moderado",
        "Repactuados + risco moderado",
        "Monitoramento",
    ),
    ("repactuados_risco_baixo", "Repactuados + risco baixo", "Monitoramento"),
    (
        "a_vencer_ate_30_dias",
        "A vencer em até 30 dias (qualquer risco)",
        "Monitoramento",
    ),
    (
        "a_vencer_mais_30_dias",
        "A vencer com mais de 30 dias (qualquer risco)",
        "Conhecimento",
    ),
]
def _clean_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _normalize_text(value: Any) -> str:
    text = "" if pd.isna(value) else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", text).strip().lower()


def _find_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    normalized = {_normalize_text(column): column for column in df.columns}
    for alias in aliases:
        found = normalized.get(_normalize_text(alias))
        if found:
            return found
    return None


def _detect_header_row(file_path, sheet_name: str) -> int:
    preview = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        nrows=20,
    )
    required = {
        _normalize_text("Ano"),
        _normalize_text("Relatório de Auditoria"),
        _normalize_text("Status"),
    }

    for row_index, row in preview.iterrows():
        values = {_normalize_text(value) for value in row if not pd.isna(value)}
        if required.issubset(values):
            return int(row_index)

    raise ValueError(
        "Não foi possível localizar automaticamente o cabeçalho da aba Dados."
    )


def _read_database() -> pd.DataFrame:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"A base de dados não foi encontrada em: {EXCEL_FILE}"
        )

    workbook = pd.ExcelFile(EXCEL_FILE)
    sheet_name = "Dados" if "Dados" in workbook.sheet_names else workbook.sheet_names[0]

    header_row = _detect_header_row(EXCEL_FILE, sheet_name)
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=header_row)
    df.columns = [_clean_header(column) for column in df.columns]

    unnamed = [column for column in df.columns if column.lower().startswith("unnamed")]
    df = df.drop(columns=unnamed, errors="ignore")
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")

    year_column = _find_column(df, ["Ano"])
    if year_column:
        df = df[df[year_column].notna()].copy()

    status_column = _find_column(df, FILTER_COLUMN_MAP["Status"])
    if status_column:
        status_labels = {
            "vencido": "Vencido",
            "solucionado": "Solucionado",
            "sem plano": "Sem Plano",
            "a vencer": "A vencer",
            "repactuado": "Repactuado",
        }
        df[status_column] = df[status_column].apply(
            lambda value: status_labels.get(_normalize_text(value), value)
            if not pd.isna(value) else value
        )

    df.reset_index(drop=True, inplace=True)
    return df


def load_data() -> pd.DataFrame:
    df = _read_database()

    missing = []
    for filter_name, aliases in FILTER_COLUMN_MAP.items():
        if _find_column(df, aliases) is None:
            missing.append(filter_name)

    if missing:
        raise ValueError(
            "A planilha não possui campos necessários para os filtros: "
            + ", ".join(missing)
        )

    return df


def load_reincidentes() -> pd.DataFrame:
    return get_recurrent_plans(load_data())


def get_recurrent_plans(df: pd.DataFrame) -> pd.DataFrame:
    reincidente_column = _find_column(
        df,
        REINCIDENTE_COLUMN_ALIASES,
    )

    if not reincidente_column:
        raise ValueError(
            "A aba Dados não possui a coluna Reincidente."
        )

    reincidentes = df[reincidente_column].map(_normalize_text)
    selected = reincidentes.isin({"sim", "s", "yes", "true", "1"})

    return df.loc[selected].reset_index(drop=True)


def get_active_plans(df: pd.DataFrame) -> pd.DataFrame:
    return df.loc[_status_series(df).isin(ACTIVE_STATUS_KEYS)].reset_index(drop=True)


def get_filter_source_columns(df: pd.DataFrame) -> dict[str, str]:
    return {
        label: _find_column(df, aliases)
        for label, aliases in FILTER_COLUMN_MAP.items()
    }


def apply_filters(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    filtered = df.copy()
    source_columns = get_filter_source_columns(df)

    for label, value in filters.items():
        source_column = source_columns.get(label)
        if value and source_column:
            filtered = filtered[
                filtered[source_column].fillna("").astype(str).str.strip() == str(value).strip()
            ]

    return filtered.reset_index(drop=True)


def get_filter_options(df: pd.DataFrame) -> dict:
    source_columns = get_filter_source_columns(df)
    result = {}

    for label in FILTER_COLUMNS:
        column = source_columns[label]
        values = df[column].dropna().astype(str).str.strip()
        values = values[values != ""]

        if label == "Ano":
            def year_sort(value: str):
                try:
                    return int(float(value))
                except ValueError:
                    return value
            result[label] = sorted(values.unique().tolist(), key=year_sort)
        else:
            result[label] = sorted(values.unique().tolist(), key=str.casefold)

    return result


def get_semaforo_filter_options(df: pd.DataFrame) -> dict:
    risk_column = _find_column(
        df,
        FILTER_COLUMN_MAP["Classificação de Risco"],
    )
    available_risks = set()
    if risk_column:
        available_risks = {
            _semaforo_risk(value)
            for value in df[risk_column]
            if _semaforo_risk(value)
        }

    return {
        "Situação": ["Atenção", "Monitoramento", "Conhecimento"],
        "Classificação de Risco": [
            risk
            for risk in ["Alto", "Moderado", "Significativo", "Baixo"]
            if risk in available_risks
        ],
        "Status": ["Não reincidente", "Reincidente"],
    }


def apply_semaforo_filters(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    filtered = df.copy()

    situation = filters.get("Situação")
    if situation:
        situations = _semaforo_rule_data(filtered)["situacoes"]
        filtered = filtered[
            situations.map(_normalize_text) == _normalize_text(situation)
        ]

    risk = filters.get("Classificação de Risco")
    if risk:
        risk_column = _find_column(
            filtered,
            FILTER_COLUMN_MAP["Classificação de Risco"],
        )
        if risk_column:
            filtered = filtered[
                filtered[risk_column].map(_semaforo_risk) == _semaforo_risk(risk)
            ]

    recurrence_status = filters.get("Status")
    if recurrence_status:
        recurrent_column = _find_column(filtered, REINCIDENTE_COLUMN_ALIASES)
        recurrent_values = (
            filtered[recurrent_column].map(_normalize_text)
            if recurrent_column
            else pd.Series("", index=filtered.index)
        )
        recurrent_mask = recurrent_values.isin({"sim", "s", "yes", "true", "1"})
        if _normalize_text(recurrence_status) == "reincidente":
            filtered = filtered[recurrent_mask]
        else:
            filtered = filtered[~recurrent_mask]

    return filtered.reset_index(drop=True)


def _status_series(df: pd.DataFrame) -> pd.Series:
    column = _find_column(df, FILTER_COLUMN_MAP["Status"])
    return df[column].map(_normalize_text) if column else pd.Series(dtype="object")


def _risk_series(df: pd.DataFrame) -> pd.Series:
    column = _find_column(df, FILTER_COLUMN_MAP["Classificação de Risco"])
    return df[column].map(_normalize_text) if column else pd.Series(dtype="object")


def _repactuation_values(df: pd.DataFrame) -> pd.Series:
    column = _find_column(df, ["Quant. Repactuações", "Quantidade de Repactuações"])
    if not column:
        return pd.Series([0] * len(df), index=df.index, dtype="float64")
    return pd.to_numeric(df[column], errors="coerce").fillna(0)


def _display_label(value: str) -> str:
    labels = {
        "alto": "Alto",
        "significativo": "Significativo",
        "moderado": "Moderado",
        "baixo": "Baixo",
        "a vencer": "A vencer",
        "repactuado": "Repactuado",
        "vencido": "Vencido",
        "solucionado": "Solucionado",
        "sem plano": "Sem Plano",
    }
    return labels.get(value, value.title() if value else "Não informado")


def _semaforo_risk(value: Any) -> str:
    normalized = _normalize_text(value)

    if "significativo" in normalized:
        return "Significativo"
    if "moderado" in normalized:
        return "Moderado"
    if "baixo" in normalized:
        return "Baixo"
    if "alto" in normalized:
        return "Alto"
    return ""


def _concise_plan_text(
    value: Any,
    maximum: int = 400,
    empty_label: str = "Plano de ação sem descrição",
) -> str:
    text = re.sub(r"\s+", " ", "" if pd.isna(value) else str(value)).strip()
    if not text:
        return empty_label
    return text if len(text) <= maximum else f"{text[:maximum - 3].rstrip()}..."


def _parse_dashboard_date(value: Any) -> pd.Timestamp:
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value.normalize()

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        parsed = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
        return parsed.normalize() if not pd.isna(parsed) else pd.NaT

    match = re.search(
        r"(?<!\d)(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})(?!\d)",
        str(value),
    )
    if match:
        day, month, year = (int(part) for part in match.groups())
        if year < 100:
            year += 2000
        try:
            return pd.Timestamp(year=year, month=month, day=day)
        except ValueError:
            return pd.NaT

    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    return parsed.normalize() if not pd.isna(parsed) else pd.NaT


def _semaforo_deadlines(df: pd.DataFrame) -> pd.Series:
    expected_column = _find_column(
        df,
        [
            "Data Prevista de Conclusão",
            "Data Prevista Conclusão",
            "Prazo Atual",
            "Data Limite",
            "Prazo",
        ],
    )

    if not expected_column:
        return pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")

    return df[expected_column].map(_parse_dashboard_date)


def _semaforo_rule_data(
    df: pd.DataFrame,
    reference_date: Any = None,
) -> dict:
    risk_column = _find_column(
        df,
        FILTER_COLUMN_MAP["Classificação de Risco"],
    )

    if not risk_column:
        raise ValueError(
            "A coluna Classificação Risco não foi encontrada."
        )

    status = _status_series(df)
    risks = df[risk_column].map(_semaforo_risk)
    deadlines = _semaforo_deadlines(df)
    reference = (
        pd.Timestamp(reference_date).normalize()
        if reference_date is not None
        else pd.Timestamp.today().normalize()
    )
    day_differences = (deadlines - reference).dt.days
    active = status.isin(ACTIVE_STATUS_KEYS)
    expired = active & (status == "vencido")
    repactuated = active & ~expired & (status == "repactuado")
    due = active & ~expired & ~repactuated & (status == "a vencer")
    rule_masks = {
        "vencidos_qualquer_risco": expired,
        "repactuados_risco_alto": repactuated & (risks == "Alto"),
        "repactuados_risco_significativo": repactuated & (risks == "Significativo"),
        "repactuados_risco_moderado": repactuated & (risks == "Moderado"),
        "repactuados_risco_baixo": repactuated & (risks == "Baixo"),
        "a_vencer_ate_30_dias": due & day_differences.notna() & (day_differences <= 30),
        "a_vencer_mais_30_dias": due & (day_differences.isna() | (day_differences > 30)),
    }
    situations = pd.Series("", index=df.index, dtype="object")
    for key, _, category in SEMAFORO_RULES:
        situations.loc[rule_masks[key]] = category

    return {
        "status": status,
        "riscos": risks,
        "prazos": deadlines,
        "data_referencia": reference,
        "diferencas_dias": day_differences,
        "ativos": active,
        "regras": rule_masks,
        "situacoes": situations,
        "a_vencer": due,
    }


def get_semaforo_indicators(
    df: pd.DataFrame,
    reference_date: Any = None,
) -> dict:
    rule_data = _semaforo_rule_data(df, reference_date)
    risks = rule_data["riscos"]
    reference = rule_data["data_referencia"]
    day_differences = rule_data["diferencas_dias"]
    active = rule_data["ativos"]
    rule_masks = rule_data["regras"]
    due = rule_data["a_vencer"]
    rule_counts = {
        key: int(mask.sum())
        for key, mask in rule_masks.items()
    }
    total_classified = sum(rule_counts.values())
    active_total = int(active.sum())
    plan_column = _find_column(df, ["Plano de ação", "Plano de Ação"])
    report_column = _find_column(
        df,
        FILTER_COLUMN_MAP["Relatório de Auditoria"],
    )
    recommendation_column = _find_column(
        df,
        ["Oportunidade de Melhoria/ Recomendação", "Recomendação"],
    )
    rules = []
    for key, condition, category in SEMAFORO_RULES:
        selected_rows = df.loc[rule_masks[key]]
        plans = []
        reports = []
        for _, row in selected_rows.iterrows():
            plan_value = row.get(plan_column, "") if plan_column else ""
            missing_plan = pd.isna(plan_value) or not re.sub(
                r"\s+", " ", str(plan_value)
            ).strip()
            if missing_plan and recommendation_column:
                plan_value = row.get(recommendation_column, "")
            plans.append(_concise_plan_text(plan_value))
            report_value = row.get(report_column, "") if report_column else ""
            reports.append(
                _concise_plan_text(
                    report_value,
                    empty_label="Relatório de Auditoria não informado",
                )
            )
        rules.append(
            {
                "chave": key,
                "condicao": condition,
                "categoria": category,
                "quantidade": rule_counts[key],
                "percentual": (
                    round(rule_counts[key] / active_total * 100, 2)
                    if active_total
                    else 0
                ),
                "planos": plans,
                "relatorios": reports,
            }
        )
    risk_order = ["Alto", "Significativo", "Moderado", "Baixo"]
    active_risks = risks[active]
    risk_distribution = {
        label: {
            "quantidade": int((active_risks == label).sum()),
            "percentual": (
                round(int((active_risks == label).sum()) / active_total * 100, 2)
                if active_total
                else 0
            ),
        }
        for label in risk_order
    }
    condition_order = ["Atenção", "Monitoramento", "Conhecimento"]
    active_counts = {
        label: sum(
            rule["quantidade"]
            for rule in rules
            if rule["categoria"] == label
        )
        for label in condition_order
    }
    cards = {
        label: {
            "quantidade": active_counts[label],
            "percentual": (
                round(active_counts[label] / active_total * 100, 2)
                if active_total
                else 0
            ),
        }
        for label in condition_order
    }

    return {
        "total_planos_ativos": active_total,
        "total_classificados": total_classified,
        "total_nao_classificados": max(0, active_total - total_classified),
        "condicoes": cards,
        "condicoes_ativas": active_counts,
        "regras": rules,
        "distribuicao_risco": risk_distribution,
        "a_vencer_sem_prazo": int((due & day_differences.isna()).sum()),
        "data_referencia": reference.strftime("%d/%m/%Y"),
    }


def get_reincidence_indicators(
    all_plans: pd.DataFrame,
    recurrent_plans: pd.DataFrame,
) -> dict:
    active_plans = get_active_plans(all_plans)
    active_recurrent_plans = get_active_plans(recurrent_plans)
    active_count = int(len(active_plans))
    recurrent_count = int(len(active_recurrent_plans))

    risk_column = _find_column(
        active_recurrent_plans,
        FILTER_COLUMN_MAP["Classificação de Risco"],
    )
    risk_order = ["Alto", "Significativo", "Moderado", "Baixo"]
    if risk_column:
        risks = active_recurrent_plans[risk_column].map(_semaforo_risk)
        criticality = {
            label: int((risks == label).sum())
            for label in risk_order
        }
    else:
        criticality = {label: 0 for label in risk_order}

    report_column = _find_column(
        active_recurrent_plans,
        FILTER_COLUMN_MAP["Relatório de Auditoria"],
    )
    reports = []
    if report_column:
        report_values = (
            active_recurrent_plans[report_column]
            .fillna("Não informado")
            .astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
            .replace("", "Não informado")
        )
        report_counts = report_values.value_counts()
        reports = [
            {
                "relatorio": str(label),
                "quantidade": int(quantity),
                "percentual": (
                    round(int(quantity) / recurrent_count * 100, 2)
                    if recurrent_count
                    else 0
                ),
            }
            for label, quantity in report_counts.items()
        ]

    return {
        "planos_monitoramento": active_count,
        "planos_reincidentes": recurrent_count,
        "taxa_reincidencia": (
            round(recurrent_count / active_count * 100, 2)
            if active_count
            else 0
        ),
        "criticidade_reincidente": criticality,
        "reincidentes_por_relatorio": reports,
    }


def _group_status_table(df: pd.DataFrame, group_label: str) -> list[dict]:
    group_column = get_filter_source_columns(df).get(group_label)
    if not group_column:
        return []

    status_order = [
        ("a vencer", "A vencer"),
        ("repactuado", "Repactuado"),
        ("vencido", "Vencido"),
        ("solucionado", "Solucionado"),
        ("sem plano", "Sem Plano"),
    ]

    working = pd.DataFrame({
        "Grupo": df[group_column].fillna("Não informado").astype(str).str.strip(),
        "StatusNorm": _status_series(df),
    })
    working["Grupo"] = working["Grupo"].replace("", "Não informado")

    if working.empty:
        return []

    pivot = pd.crosstab(working["Grupo"], working["StatusNorm"])
    grand_total = int(len(working))
    rows = []

    for group, values in pivot.iterrows():
        row = {group_label: group or "Não informado"}
        total = 0
        for normalized, label in status_order:
            quantity = int(values.get(normalized, 0))
            row[label] = quantity
            total += quantity
        row["Total"] = total
        row["%"] = round((total / grand_total * 100), 2) if grand_total else 0
        rows.append(row)

    return sorted(rows, key=lambda row: (-row["Total"], str(row[group_label]).casefold()))



def _origin_overview(df: pd.DataFrame) -> dict:
    origin_column = _find_column(df, FILTER_COLUMN_MAP["Origem"])
    labels = [
        "Auditoria Interna",
        "Auditoria Externa",
        "Auditoria Independente",
        "TCU/CGU",
    ]
    result = {
        label: {"quantidade": 0, "percentual": 0}
        for label in labels
    }

    if not origin_column:
        return result

    origins = df[origin_column].map(_normalize_text)
    total = int(len(df))

    for value in origins:
        if "auditoria independente" in value:
            label = "Auditoria Independente"
        elif "auditoria interna" in value:
            label = "Auditoria Interna"
        elif "auditoria externa" in value:
            label = "Auditoria Externa"
        elif "tcu" in value or "cgu" in value:
            label = "TCU/CGU"
        else:
            continue
        result[label]["quantidade"] += 1

    for label in labels:
        quantity = result[label]["quantidade"]
        result[label]["percentual"] = (
            round((quantity / total * 100), 1) if total else 0
        )

    return result



def _initial_completion_periods(df: pd.DataFrame) -> list[dict]:
    date_column = _find_column(
        df,
        [
            "Data Prevista de Conclusão",
            "Data Prevista Conclusão",
            "Data Inicial de Conclusão",
            "Conclusão Inicial",
        ],
    )

    date_pattern = re.compile(
        r"(?<!\d)(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})(?!\d)"
    )
    events: list[tuple[int, int]] = []

    if date_column:
        status = _status_series(df)
        selected_dates = df.loc[
            status.isin(["vencido", "repactuado"]),
            date_column,
        ]

        for value in selected_dates.dropna():
            if isinstance(value, pd.Timestamp):
                events.append((int(value.year), int(value.month)))
                continue

            match = date_pattern.search(str(value))
            if match:
                _day, month, year = match.groups()
                year_number = int(year)
                if year_number < 100:
                    year_number += 2000

                month_number = int(month)
                if 1 <= month_number <= 12:
                    events.append((year_number, month_number))
                continue

            parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
            if not pd.isna(parsed):
                events.append((int(parsed.year), int(parsed.month)))

    final_year = max([2026, *[year for year, _month in events]])

    periods = [
        (year, start, 6 if start == 1 else 12)
        for year in range(2023, final_year + 1)
        for start in (1, 7)
    ]
    counts = {(year, start): 0 for year, start, _end in periods}

    for year, month in events:
        semester_start = 1 if month <= 6 else 7
        key = (year, semester_start)
        if key in counts:
            counts[key] += 1

    return [
        {
            "periodo": f"{start:02d}/{year} a {end:02d}/{year}",
            "quantidade": counts[(year, start)],
        }
        for year, start, end in periods
    ]

def get_indicators(df: pd.DataFrame) -> dict:
    status = _status_series(df)
    risks = _risk_series(df)
    repactuations = _repactuation_values(df)

    active_mask = status.isin(["a vencer", "repactuado", "vencido"])
    active_count = int(active_mask.sum())
    expired_count = int((status == "vencido").sum())
    repactuated_mask = status == "repactuado"
    repactuated_plans = int(repactuated_mask.sum())
    repactuation_total = int(repactuations[repactuated_mask].sum())

    risk_counts = {}
    for risk in ["alto", "significativo", "moderado", "baixo"]:
        risk_counts[_display_label(risk)] = int((active_mask & (risks == risk)).sum())

    distribution = {
        "1": int(((repactuations == 1) & repactuated_mask).sum()),
        "2": int(((repactuations == 2) & repactuated_mask).sum()),
        "3": int(((repactuations == 3) & repactuated_mask).sum()),
        "Acima de 3": int(((repactuations > 3) & repactuated_mask).sum()),
    }

    total_status = int(len(df))
    status_overview = {}
    for normalized, label in [
        ("solucionado", "Solucionado"),
        ("repactuado", "Repactuado"),
        ("a vencer", "A vencer"),
        ("vencido", "Vencido"),
        ("sem plano", "Sem Plano"),
    ]:
        quantity = int((status == normalized).sum())
        status_overview[label] = {
            "quantidade": quantity,
            "percentual": round((quantity / total_status * 100), 1) if total_status else 0,
        }

    solved_count = int((status == "solucionado").sum())
    completion_periods = _initial_completion_periods(df)

    return {
        "total_registros": int(len(df)),
        "total_planos_acao": int(len(df)),
        "planos_solucionados": solved_count,
        "planos_ativos": active_count,
        "planos_vencidos": expired_count,
        "planos_repactuados": repactuated_plans,
        "quantidade_repactuacoes": repactuation_total,
        "taxa_repactuacao": round((repactuated_plans / active_count * 100), 2) if active_count else 0,
        "criticidade_riscos": risk_counts,
        "distribuicao_repactuacoes": distribution,
        "planos_por_diretoria": _group_status_table(df, "Diretoria"),
        "planos_por_unidade": _group_status_table(df, "Unidade/Agência"),
        "periodos_conclusao_planos": completion_periods,
        "repactuacoes_por_periodo": completion_periods,
        "visao_geral_status": status_overview,
        "visao_geral_origem": _origin_overview(df),
    }


def get_information_summary(df: pd.DataFrame) -> dict:
    source_columns = get_filter_source_columns(df)
    year_column = source_columns.get("Ano")

    years = []
    if year_column:
        year_values = pd.to_numeric(df[year_column], errors="coerce").dropna()
        years = sorted({int(value) for value in year_values})

    if not years:
        period = "Não informado"
    elif len(years) == 1:
        period = str(years[0])
    else:
        period = f"{years[0]} a {years[-1]}"

    return {
        "arquivo": EXCEL_FILE.name,
        "periodo": period,
    }


def _json_value(value: Any) -> Any:
    if value is None or (isinstance(value, float) and math.isnan(value)) or pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%d/%m/%Y")
    if hasattr(value, "item"):
        try:
            return value.item()
        except (ValueError, AttributeError):
            pass
    return value


def dataframe_records(df: pd.DataFrame) -> list[dict]:
    return [
        {column: _json_value(value) for column, value in row.items()}
        for row in df.to_dict(orient="records")
    ]
