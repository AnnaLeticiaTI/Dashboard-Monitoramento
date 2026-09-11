from flask import Blueprint, jsonify, request, send_file

from services.data_service import (
    FILTER_COLUMNS,
    SEMAFORO_FILTER_COLUMNS,
    apply_filters,
    apply_semaforo_filters,
    dataframe_records,
    get_active_plans,
    get_filter_options,
    get_indicators,
    get_recurrent_plans,
    get_reincidence_indicators,
    get_semaforo_filter_options,
    get_semaforo_indicators,
    load_data,
)
from services.carteira_service import (
    CARTEIRA_FILTER_COLUMNS,
    apply_aguardando_filters,
    apply_carteira_filters,
    carteira_dataframe_records,
    get_carteira_filter_options,
    get_carteira_indicators,
    get_carteira_source_summary,
    load_aguardando_pa_data,
    load_carteira_data,
)
from services.export_service import (
    export_carteira_excel,
    export_carteira_pdf,
    export_carteira_table_excel,
    export_carteira_table_pdf,
    export_indicators_excel,
    export_indicators_pdf,
    export_paint_excel,
    export_paint_pdf,
    export_paint_table_excel,
    export_paint_table_pdf,
    export_semaforo_excel,
    export_semaforo_pdf,
    export_table_excel,
    export_table_pdf,
)
from services.paint_service import (
    PAINT_FILTER_COLUMNS,
    apply_paint_filters,
    get_paint_filter_options,
    get_paint_indicators,
    get_paint_source_summary,
    load_paint_data,
    paint_dataframe_records,
)

api = Blueprint("api", __name__)
def _filters_from_request() -> dict:
    return {
        key: value
        for key, value in request.args.items()
        if key in FILTER_COLUMNS and value
    }


def _semaforo_filters_from_request() -> dict:
    return {
        key: value
        for key, value in request.args.items()
        if key in SEMAFORO_FILTER_COLUMNS and value
    }


def _paint_filters_from_request() -> dict:
    return {
        key: value
        for key, value in request.args.items()
        if key in PAINT_FILTER_COLUMNS and value
    }


def _carteira_filters_from_request() -> dict:
    return {
        key: value
        for key, value in request.args.items()
        if key in CARTEIRA_FILTER_COLUMNS and value
    }


def _paint_export_filters(filters: dict) -> dict:
    return {
        "Processo(s)" if key == "Processo" else key: value
        for key, value in filters.items()
    }


@api.route("/api/dashboard")
def dashboard_data():
    df = load_data()
    filters = _filters_from_request()
    filtered = apply_filters(df, filters)

    return jsonify({
        "indicators": get_indicators(filtered),
        "filters": get_filter_options(df),
        "filter_columns": FILTER_COLUMNS,
        "table_columns": filtered.columns.tolist(),
        "active_filters": filters,
        "rows": dataframe_records(filtered),
    })


@api.route("/api/semaforo")
def semaforo_data():
    try:
        all_plans = load_data()
        all_active_plans = get_active_plans(all_plans)
        filters = _semaforo_filters_from_request()
        filtered_plans = apply_semaforo_filters(all_plans, filters)
        active_plans = get_active_plans(filtered_plans)
        recurrent_plans = get_recurrent_plans(active_plans)
        indicators = get_semaforo_indicators(active_plans)
        indicators["reincidencia"] = get_reincidence_indicators(
            active_plans,
            recurrent_plans,
        )

        return jsonify({
            "indicators": indicators,
            "filters": get_semaforo_filter_options(all_active_plans),
            "filter_columns": SEMAFORO_FILTER_COLUMNS,
            "active_filters": filters,
            "table_columns": active_plans.columns.tolist(),
            "rows": dataframe_records(active_plans),
        })
    except (FileNotFoundError, ValueError) as error:
        return jsonify({
            "error": str(error),
        }), 400


@api.route("/api/paint")
def paint_data():
    try:
        all_rows = load_paint_data()
        filters = _paint_filters_from_request()
        filtered_rows = apply_paint_filters(all_rows, filters)

        return jsonify({
            "indicators": get_paint_indicators(filtered_rows),
            "filters": get_paint_filter_options(all_rows),
            "filter_columns": PAINT_FILTER_COLUMNS,
            "active_filters": filters,
            "table_columns": filtered_rows.columns.tolist(),
            "rows": paint_dataframe_records(filtered_rows),
            "source": get_paint_source_summary(all_rows),
        })
    except (FileNotFoundError, ValueError) as error:
        return jsonify({"error": str(error)}), 400


@api.route("/api/carteira")
def carteira_data():
    try:
        all_carteira = load_carteira_data()
        all_aguardando = load_aguardando_pa_data()
        filters = _carteira_filters_from_request()
        filtered_carteira = apply_carteira_filters(all_carteira, filters)
        filtered_aguardando = apply_aguardando_filters(
            all_aguardando,
            filtered_carteira,
        )

        return jsonify({
            "indicators": get_carteira_indicators(filtered_carteira),
            "filters": get_carteira_filter_options(all_carteira),
            "filter_columns": CARTEIRA_FILTER_COLUMNS,
            "active_filters": filters,
            "tables": {
                "carteira": {
                    "columns": filtered_carteira.columns.tolist(),
                    "rows": carteira_dataframe_records(filtered_carteira),
                },
                "aguardando": {
                    "columns": filtered_aguardando.columns.tolist(),
                    "rows": carteira_dataframe_records(filtered_aguardando),
                },
            },
            "source": get_carteira_source_summary(all_carteira),
        })
    except (FileNotFoundError, ValueError) as error:
        return jsonify({"error": str(error)}), 400


@api.route("/export/indicators/<file_type>")
def export_indicators(file_type):
    filters = _filters_from_request()
    df = apply_filters(load_data(), filters)
    indicators = get_indicators(df)

    if file_type == "pdf":
        file_path = export_indicators_pdf(indicators, filters)
    elif file_type == "excel":
        file_path = export_indicators_excel(indicators, filters)
    else:
        return jsonify({"error": "Formato inválido."}), 400

    return send_file(file_path, as_attachment=True)


@api.route("/export/semaforo/<file_type>")
def export_semaforo(file_type):
    all_plans = load_data()
    filters = _semaforo_filters_from_request()
    filtered_plans = apply_semaforo_filters(all_plans, filters)
    active_plans = get_active_plans(filtered_plans)
    recurrent_plans = get_recurrent_plans(active_plans)
    indicators = get_semaforo_indicators(active_plans)
    indicators["reincidencia"] = get_reincidence_indicators(
        active_plans,
        recurrent_plans,
    )

    if file_type == "pdf":
        file_path = export_semaforo_pdf(indicators, active_plans, filters)
    elif file_type == "excel":
        file_path = export_semaforo_excel(indicators, active_plans, filters)
    else:
        return jsonify({"error": "Formato inválido."}), 400

    return send_file(file_path, as_attachment=True)


@api.route("/export/semaforo/table/<file_type>")
def export_semaforo_table(file_type):
    semaforo_filters = _semaforo_filters_from_request()
    active_plans = apply_semaforo_filters(
        get_active_plans(load_data()),
        semaforo_filters,
    )
    filters = {"Status dos planos": "Planos ativos", **semaforo_filters}

    if file_type == "excel":
        file_path = export_table_excel(active_plans, filters)
    elif file_type == "pdf":
        file_path = export_table_pdf(active_plans, filters)
    else:
        return jsonify({"error": "Formato inválido."}), 400

    return send_file(file_path, as_attachment=True)


@api.route("/export/paint/<file_type>")
def export_paint(file_type):
    all_rows = load_paint_data()
    filters = _paint_filters_from_request()
    filtered_rows = apply_paint_filters(all_rows, filters)
    indicators = get_paint_indicators(filtered_rows)
    source = get_paint_source_summary(all_rows)
    export_filters = _paint_export_filters(filters)

    if file_type == "pdf":
        file_path = export_paint_pdf(indicators, filtered_rows, export_filters, source)
    elif file_type == "excel":
        file_path = export_paint_excel(indicators, filtered_rows, export_filters, source)
    else:
        return jsonify({"error": "Formato inválido."}), 400

    return send_file(file_path, as_attachment=True)


@api.route("/export/paint/table/<file_type>")
def export_paint_table(file_type):
    filters = _paint_filters_from_request()
    filtered_rows = apply_paint_filters(load_paint_data(), filters)
    export_filters = _paint_export_filters(filters)

    if file_type == "excel":
        file_path = export_paint_table_excel(filtered_rows, export_filters)
    elif file_type == "pdf":
        file_path = export_paint_table_pdf(filtered_rows, export_filters)
    else:
        return jsonify({"error": "Formato inválido."}), 400

    return send_file(file_path, as_attachment=True)


@api.route("/export/carteira/<file_type>")
def export_carteira(file_type):
    all_carteira = load_carteira_data()
    filters = _carteira_filters_from_request()
    filtered_carteira = apply_carteira_filters(all_carteira, filters)
    filtered_aguardando = apply_aguardando_filters(
        load_aguardando_pa_data(),
        filtered_carteira,
    )
    indicators = get_carteira_indicators(filtered_carteira)
    source = get_carteira_source_summary(all_carteira)

    if file_type == "pdf":
        file_path = export_carteira_pdf(
            indicators,
            filtered_carteira,
            filtered_aguardando,
            filters,
            source,
        )
    elif file_type == "excel":
        file_path = export_carteira_excel(
            indicators,
            filtered_carteira,
            filtered_aguardando,
            filters,
            source,
        )
    else:
        return jsonify({"error": "Formato inválido."}), 400

    return send_file(file_path, as_attachment=True)


@api.route("/export/carteira/table/<file_type>")
def export_carteira_table(file_type):
    filters = _carteira_filters_from_request()
    table_key = request.args.get("table", "carteira")
    filtered_carteira = apply_carteira_filters(load_carteira_data(), filters)
    if table_key == "aguardando":
        dataframe = apply_aguardando_filters(
            load_aguardando_pa_data(),
            filtered_carteira,
        )
    else:
        table_key = "carteira"
        dataframe = filtered_carteira

    if file_type == "pdf":
        file_path = export_carteira_table_pdf(dataframe, filters, table_key)
    elif file_type == "excel":
        file_path = export_carteira_table_excel(dataframe, filters, table_key)
    else:
        return jsonify({"error": "Formato inválido."}), 400

    return send_file(file_path, as_attachment=True)


@api.route("/export/table/<file_type>")
def export_table(file_type):
    filters = _filters_from_request()
    df = apply_filters(load_data(), filters)

    if file_type == "excel":
        file_path = export_table_excel(df, filters)
    elif file_type == "pdf":
        file_path = export_table_pdf(df, filters)
    else:
        return jsonify({"error": "Formato inválido."}), 400

    return send_file(file_path, as_attachment=True)
