let dashboardRows = [];
let tableColumns = [];
let semaforoRows = [];
let semaforoTableColumns = [];
let repactuationChart = null;
let activePlansChart = null;
let directorateChart = null;
let unitChart = null;
let statusOverviewChart = null;
let originOverviewChart = null;
let repactuationPeriodChart = null;
let semaforoRiskChart = null;
let semaforoRiskDistributionChart = null;
let recurrenceCriticalityChart = null;
let paintRows = [];
let paintTableColumns = [];
let paintProcessStatusChart = null;
let paintStatusDistributionChart = null;
let carteiraTables = {
    carteira: { columns: [], rows: [] },
    aguardando: { columns: [], rows: [] }
};
let carteiraActiveTable = "carteira";
let carteiraStatusChart = null;
let carteiraPaintChart = null;


const legendWhiteBackground = {
    id: "legendWhiteBackground",
    beforeDraw(chart) {
        const legend = chart.legend;
        if (!legend || legend.options.display === false || legend.position !== "bottom") return;

        const ctx = chart.ctx;
        const paddingX = 12;
        const paddingY = 7;
        const x = legend.left - paddingX;
        const y = legend.top - paddingY;
        const width = legend.width + paddingX * 2;
        const height = legend.height + paddingY * 2;

        ctx.save();
        ctx.fillStyle = "rgba(255, 255, 255, .96)";
        ctx.strokeStyle = "rgba(36, 54, 75, .12)";
        ctx.lineWidth = 1;
        ctx.beginPath();
        if (typeof ctx.roundRect === "function") {
            ctx.roundRect(x, y, width, height, 10);
        } else {
            ctx.rect(x, y, width, height);
        }
        ctx.fill();
        ctx.stroke();
        ctx.restore();
    }
};

const filterColumns = [
    "Origem",
    "Relatório de Auditoria",
    "Diretoria",
    "Unidade/Agência",
    "Responsável",
    "Classificação de Risco",
    "Status",
    "Ano"
];

const semaforoFilterColumns = [
    "Situação",
    "Classificação de Risco",
    "Status"
];

const paintFilterColumns = [
    "Ano",
    "Auditoria",
    "Processo",
    "Realização",
    "Status"
];

const paintFilterLabels = {
    Processo: "Processo(s)"
};

const carteiraFilterColumns = ["Auditoria", "PAINT", "Status"];

function currentFilterParams() {
    const params = new URLSearchParams();
    document.querySelectorAll(".filter-select").forEach(select => {
        if (select.value) params.append(select.name, select.value);
    });
    return params;
}

function currentSemaforoFilterParams() {
    const params = new URLSearchParams();
    document.querySelectorAll(".semaforo-filter-select").forEach(select => {
        if (select.value) params.append(select.name, select.value);
    });
    return params;
}

function currentPaintFilterParams() {
    const params = new URLSearchParams();
    document.querySelectorAll(".paint-filter-select").forEach(select => {
        if (select.value) params.append(select.name, select.value);
    });
    return params;
}

function currentCarteiraFilterParams() {
    const params = new URLSearchParams();
    document.querySelectorAll(".carteira-filter-select").forEach(select => {
        if (select.value) params.append(select.name, select.value);
    });
    return params;
}

async function loadDashboard() {
    if (!document.getElementById("filterBar")) return;

    try {
        const response = await fetch(`/api/dashboard?${currentFilterParams().toString()}`);
        const data = await response.json();

        if (!response.ok) throw new Error(data.error || "Não foi possível carregar a base.");

        dashboardRows = data.rows;
        tableColumns = data.table_columns;

        renderFilters(data.filters);
        renderActiveFilters(data.active_filters);
        renderCards(data.indicators);
        renderStatusOverview(data.indicators);
        renderOriginOverview(data.indicators);
        renderRepactuationKpis(data.indicators);
        renderRepactuationPeriod(data.indicators);
        renderTable(dashboardRows);
        renderCharts(data.indicators);
        renderGroupedTable("directorateIndicatorTable", data.indicators.planos_por_diretoria, "Diretoria");
        renderGroupedTable("unitIndicatorTable", data.indicators.planos_por_unidade, "Unidade/Agência");
    } catch (error) {
        console.error(error);
        document.getElementById("summaryCards").innerHTML = `
            <div class="dashboard-error">${error.message}</div>
        `;
    }
}

function renderFilters(filters) {
    const filterBar = document.getElementById("filterBar");
    const oldValues = {};

    document.querySelectorAll(".filter-select").forEach(select => {
        oldValues[select.name] = select.value;
    });

    filterBar.innerHTML = "";

    filterColumns.forEach(column => {
        const select = document.createElement("select");
        select.name = column;
        select.className = "filter-select";
        select.innerHTML = `<option value="">${column}</option>`;

        (filters[column] || []).forEach(option => {
            const item = document.createElement("option");
            item.value = option;
            item.textContent = option;
            select.appendChild(item);
        });

        select.value = oldValues[column] || "";
        select.addEventListener("change", loadDashboard);
        filterBar.appendChild(select);
    });
}

function renderActiveFilters(activeFilters) {
    const active = Object.entries(activeFilters);
    document.getElementById("activeFilters").textContent = active.length
        ? active.map(([key, value]) => `${key}: ${value}`).join(" | ")
        : "Nenhum filtro ativo.";
}

function clearFilters() {
    document.querySelectorAll(".filter-select").forEach(select => {
        select.value = "";
    });

    const search = document.getElementById("tableSearch");
    if (search) search.value = "";

    loadDashboard();
}

function renderCards(indicators) {
    const cards = [
        { label: "Total de Planos de Ação", value: indicators.total_planos_acao, className: "kpi-total" },
        { label: "Planos Solucionados", value: indicators.planos_solucionados, className: "kpi-solved" },
        { label: "Planos Vencidos", value: indicators.planos_vencidos, className: "kpi-expired" }
    ];

    document.getElementById("summaryCards").innerHTML = cards.map(card => `
        <div class="indicator-card executive-kpi ${card.className}">
            <span>${card.label}</span>
            <strong>${card.value}</strong>
        </div>
    `).join("");
}

function percentageFormatter(value, context) {
    const values = context.chart.data.datasets[0].data.map(Number);
    const total = values.reduce((sum, item) => sum + item, 0);
    if (Number(value) <= 0 || total <= 0) return "";
    return `${(Number(value) / total * 100).toFixed(1).replace(".", ",")}%`;
}

function overviewChartOptions() {
    return {
        responsive: true,
        maintainAspectRatio: false,
        layout: {
            padding: { top: 8, right: 8, bottom: 18, left: 8 }
        },
        plugins: {
            legend: {
                position: "bottom",
                fullSize: true,
                labels: {
                    color: "#111111",
                    boxWidth: 11,
                    padding: 12,
                    font: { size: 10 }
                }
            },
            tooltip: { enabled: true },
            datalabels: {
                formatter: percentageFormatter,
                color: "#111111",
                backgroundColor: null,
                borderRadius: 0,
                padding: 0,
                textStrokeWidth: 0,
                font: { weight: "bold", size: 12 },
                anchor: "center",
                align: "center",
                clamp: true
            }
        }
    };
}

function renderStatusOverview(indicators) {
    const overview = indicators.visao_geral_status || {};
    const order = ["Solucionado", "Repactuado", "A vencer", "Vencido", "Sem Plano"];
    const colors = {
        "Solucionado": "#4F81BD", "Repactuado": "#FFF200", "A vencer": "#00B050",
        "Vencido": "#FF0000", "Sem Plano": "#BFBFBF"
    };
    const total = order.reduce((sum, label) => sum + Number(overview[label]?.quantidade || 0), 0);
    const tbody = document.querySelector("#statusOverviewTable tbody");
    if (tbody) {
        tbody.innerHTML = order.map(label => `
            <tr><th style="--status-color:${colors[label]}">${label}</th><td>${overview[label]?.quantidade || 0}</td></tr>
        `).join("") + `<tr class="overview-total-row"><th>TOTAL</th><td>${total}</td></tr>`;
    }

    const canvas = document.getElementById("statusOverviewChart");
    if (!canvas) return;
    if (statusOverviewChart) statusOverviewChart.destroy();
    statusOverviewChart = new Chart(canvas, {
        type: "pie",
        plugins: [ChartDataLabels, legendWhiteBackground],
        data: {
            labels: order,
            datasets: [{ data: order.map(label => Number(overview[label]?.quantidade || 0)), backgroundColor: order.map(label => colors[label]), borderColor: "#FFFFFF", borderWidth: 2 }]
        },
        options: overviewChartOptions()
    });
}

function renderOriginOverview(indicators) {
    const overview = indicators.visao_geral_origem || {};
    const order = ["Auditoria Interna", "Auditoria Externa", "Auditoria Independente", "TCU/CGU"];
    const colors = {
        "Auditoria Interna": "#7B337E", "Auditoria Externa": "#CA043E",
        "Auditoria Independente": "#FD7600", "TCU/CGU": "#FED000"
    };
    const total = order.reduce((sum, label) => sum + Number(overview[label]?.quantidade || 0), 0);
    const tbody = document.querySelector("#originOverviewTable tbody");
    if (tbody) {
        tbody.innerHTML = order.map(label => `
            <tr><th style="--status-color:${colors[label]}">${label}</th><td>${overview[label]?.quantidade || 0}</td></tr>
        `).join("") + `<tr class="overview-total-row"><th>TOTAL</th><td>${total}</td></tr>`;
    }

    const canvas = document.getElementById("originOverviewChart");
    if (!canvas) return;
    if (originOverviewChart) originOverviewChart.destroy();
    originOverviewChart = new Chart(canvas, {
        type: "pie",
        plugins: [ChartDataLabels, legendWhiteBackground],
        data: {
            labels: order,
            datasets: [{ data: order.map(label => Number(overview[label]?.quantidade || 0)), backgroundColor: order.map(label => colors[label]), borderColor: "#FFFFFF", borderWidth: 2 }]
        },
        options: overviewChartOptions()
    });
}

function renderTable(rows) {
    const table = document.getElementById("dataTable");
    table.querySelector("thead").innerHTML = `
        <tr>${tableColumns.map(column => `<th>${escapeHtml(column)}</th>`).join("")}</tr>
    `;
    table.querySelector("tbody").innerHTML = rows.map(row => `
        <tr>${tableColumns.map(column => `<td>${escapeHtml(row[column] ?? "")}</td>`).join("")}</tr>
    `).join("");
}

function formatPercent(value) {
    return `${Number(value || 0).toFixed(2).replace(".", ",")}%`;
}

function renderGroupedTable(tableId, rows, groupColumn) {
    const table = document.getElementById(tableId);
    if (!table) return;

    const columns = [groupColumn, "A vencer", "Repactuado", "Vencido", "Solucionado", "Sem Plano", "Total", "%"];
    const statusClasses = {
        "A vencer": "status-a-vencer",
        "Repactuado": "status-repactuado",
        "Vencido": "status-vencido",
        "Solucionado": "status-solucionado",
        "Sem Plano": "status-sem-plano"
    };

    table.querySelector("thead").innerHTML = `
        <tr>${columns.map(column => `<th class="${statusClasses[column] || ""}">${column}</th>`).join("")}</tr>
    `;

    const totals = rows.reduce((acc, row) => {
        ["A vencer", "Repactuado", "Vencido", "Solucionado", "Sem Plano", "Total"].forEach(column => {
            acc[column] += Number(row[column] || 0);
        });
        return acc;
    }, { "A vencer": 0, "Repactuado": 0, "Vencido": 0, "Solucionado": 0, "Sem Plano": 0, "Total": 0 });

    const dataRows = rows.map(row => `
        <tr>
            ${columns.map(column => {
                const value = column === "%" ? formatPercent(row[column]) : escapeHtml(row[column] ?? 0);
                return `<td>${value}</td>`;
            }).join("")}
        </tr>
    `).join("");

    const totalRow = `
        <tr class="grouped-total-row">
            <td>TOTAL</td>
            <td>${totals["A vencer"]}</td>
            <td>${totals.Repactuado}</td>
            <td>${totals.Vencido}</td>
            <td>${totals.Solucionado}</td>
            <td>${totals["Sem Plano"]}</td>
            <td>${totals.Total}</td>
            <td>${totals.Total ? "100,00%" : "0,00%"}</td>
        </tr>
    `;

    table.querySelector("tbody").innerHTML = dataRows + totalRow;
}


function valueWithPercentageFormatter(value, context) {
    const values = context.chart.data.datasets[0].data.map(Number);
    const total = values.reduce((sum, item) => sum + item, 0);
    const numericValue = Number(value || 0);
    if (numericValue <= 0 || total <= 0) return "";

    const percentage = (numericValue / total * 100)
        .toFixed(1)
        .replace(".", ",");
    return `${numericValue} (${percentage}%)`;
}

function chartDefaults({ showValues = false, valueFormatter = null } = {}) {
    return {
        responsive: true,
        maintainAspectRatio: false,
        layout: {
            padding: { top: 14, right: 18, bottom: 10, left: 18 }
        },
        plugins: {
            legend: {
                position: "bottom",
                labels: {
                    color: "#111111",
                    usePointStyle: true,
                    boxWidth: 10,
                    padding: 14,
                    font: { size: 10, weight: "600" }
                }
            },
            tooltip: {
                enabled: true,
                displayColors: true,
                backgroundColor: "rgba(255, 255, 255, .98)",
                titleColor: "#0E2146",
                bodyColor: "#0E2146",
                borderColor: "rgba(36, 54, 75, .24)",
                borderWidth: 1,
                padding: 11
            },
            datalabels: {
                display: showValues,
                color: "#111111",
                backgroundColor: null,
                textStrokeWidth: 0,
                font: { weight: "bold", size: 12 },
                formatter: valueFormatter || (value => Number(value) > 0 ? Number(value) : ""),
                anchor: "center",
                align: "center",
                clamp: true
            }
        }
    };
}

function groupedChartOptions(maxValue = undefined, hideXAxisNumbers = false) {
    return {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: "y",
        interaction: { mode: "index", intersect: true, axis: "y" },
        layout: {
            padding: { top: 12, right: 34, bottom: 10, left: 12 }
        },
        plugins: {
            legend: {
                position: "bottom",
                labels: {
                    color: "#111111",
                    usePointStyle: true,
                    boxWidth: 10,
                    padding: 14,
                    font: { size: 11, weight: "600" }
                }
            },
            tooltip: {
                enabled: true,
                mode: "index",
                intersect: true,
                axis: "y",
                displayColors: true,
                backgroundColor: "rgba(255, 255, 255, .98)",
                titleColor: "#0E2146",
                bodyColor: "#0E2146",
                borderColor: "rgba(36, 54, 75, .30)",
                borderWidth: 1,
                padding: 12,
                callbacks: {
                    title: items => items.length ? items[0].label : "",
                    label: context => `${context.dataset.label}: ${Number(context.raw || 0)}`
                }
            },
            datalabels: {
                color: "#111111",
                backgroundColor: null,
                textStrokeWidth: 0,
                font: { weight: "bold", size: 11 },
                formatter: value => Number(value) > 0 ? Number(value) : "",
                clamp: true
            }
        },
        scales: {
            x: {
                stacked: true,
                beginAtZero: true,
                ticks: {
                    display: !hideXAxisNumbers,
                    color: "#111111",
                    precision: 0,
                    stepSize: 1
                },
                suggestedMax: maxValue,
                grid: { color: "rgba(36, 54, 75, .12)" },
                border: { color: "rgba(36, 54, 75, .24)" },
                title: {
                    display: true,
                    text: "Quantidade de planos",
                    color: "#111111",
                    font: { weight: "700" }
                }
            },
            y: {
                stacked: true,
                grid: { display: false },
                border: { color: "rgba(36, 54, 75, .24)" },
                ticks: {
                    color: "#111111",
                    autoSkip: false,
                    font: { size: 11, weight: "700" }
                }
            }
        }
    };
}

function buildGroupedDatasets(rows) {
    const statuses = [
        { label: "A vencer", color: "#00B050" },
        { label: "Repactuado", color: "#FFF200" },
        { label: "Vencido", color: "#FF0000" },
        { label: "Solucionado", color: "#4F81BD" },
        { label: "Sem Plano", color: "#BFBFBF" }
    ];

    return statuses.map(status => ({
        label: status.label,
        data: rows.map(row => Number(row[status.label] || 0)),
        backgroundColor: status.color,
        borderColor: "rgba(255, 255, 255, .92)",
        borderWidth: 1,
        borderRadius: 8,
        borderSkipped: false
    }));
}

function renderRepactuationSummary(indicators) {
    const table = document.getElementById("repactuationSummaryTable");
    if (!table) return;

    const distribution = indicators.distribuicao_repactuacoes || {};
    const rows = [
        { label: "1", value: Number(distribution["1"] || 0), className: "repact-one" },
        { label: "2", value: Number(distribution["2"] || 0), className: "repact-two" },
        { label: "3", value: Number(distribution["3"] || 0), className: "repact-three" },
        { label: "Acima 3", value: Number(distribution["Acima de 3"] || 0), className: "repact-above" }
    ];
    const total = rows.reduce((sum, row) => sum + row.value, 0);

    table.querySelector("tbody").innerHTML = rows.map(row => {
        const percentage = total ? (row.value / total * 100) : 0;
        return `
            <tr class="${row.className}">
                <th>${row.label}</th>
                <td>${row.value}</td>
                <td>${percentage.toFixed(1).replace(".", ",")}</td>
            </tr>
        `;
    }).join("") + `
        <tr class="repact-total">
            <th>TOTAL</th>
            <td>${total}</td>
            <td>${total ? "100,0" : "0,0"}</td>
        </tr>
    `;
}


function renderRepactuationKpis(indicators) {
    const totalElement = document.getElementById("repactuatedPlansKpi");
    const rateElement = document.getElementById("repactuationRateKpi");

    if (totalElement) {
        totalElement.textContent = Number(indicators.planos_repactuados || 0);
    }

    if (rateElement) {
        rateElement.textContent = `${Number(indicators.taxa_repactuacao || 0)
            .toFixed(2)
            .replace(".", ",")}%`;
    }
}

function renderRepactuationPeriod(indicators) {
    const rows = indicators.periodos_conclusao_planos
        || indicators.repactuacoes_por_periodo
        || [];
    const table = document.getElementById("repactuationPeriodTable");

    if (table) {
        const tbody = table.querySelector("tbody");
        tbody.innerHTML = rows.map(row => `
            <tr>
                <td>${escapeHtml(row.periodo)}</td>
                <td>${Number(row.quantidade || 0)}</td>
            </tr>
        `).join("");
    }

    const canvas = document.getElementById("repactuationPeriodChart");
    if (!canvas) return;

    if (repactuationPeriodChart) repactuationPeriodChart.destroy();

    repactuationPeriodChart = new Chart(canvas, {
        type: "bar",
        plugins: [ChartDataLabels],
        data: {
            labels: rows.map(row => row.periodo),
            datasets: [{
                label: "Planos vencidos e repactuados",
                data: rows.map(row => Number(row.quantidade || 0)),
                backgroundColor: "#4F81BD",
                borderColor: "#345998",
                borderWidth: 1,
                borderRadius: 7,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: "y",
            plugins: {
                legend: { display: false },
                tooltip: { enabled: true },
                datalabels: {
                    display: false
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: { precision: 0, stepSize: 1, color: "#111111" },
                    grid: { color: "rgba(36, 54, 75, .12)" },
                    title: {
                        display: true,
                        text: "Quantidade de planos",
                        color: "#111111",
                        font: { weight: "700" }
                    }
                },
                y: {
                    grid: { display: false },
                    ticks: { color: "#111111", font: { size: 11, weight: "700" } }
                }
            }
        }
    });
}

function renderCharts(indicators) {
    const directorateRows = indicators.planos_por_diretoria || [];
    const unitRows = indicators.planos_por_unidade || [];

    renderRepactuationSummary(indicators);


    const activePlansCanvas = document.getElementById("activePlansChart");
    if (activePlansCanvas) {
        if (activePlansChart) activePlansChart.destroy();
        activePlansChart = new Chart(activePlansCanvas, {
        type: "doughnut",
        plugins: [ChartDataLabels, legendWhiteBackground],
        data: {
            labels: ["Planos repactuados", "Demais planos ativos"],
            datasets: [{
                data: [
                    Number(indicators.planos_repactuados || 0),
                    Math.max(0, Number(indicators.planos_ativos || 0) - Number(indicators.planos_repactuados || 0))
                ],
                backgroundColor: ["#f5f506", "#1565C0"],
                borderColor: "#ffffff",
                borderWidth: 2
            }]
        },
        options: chartDefaults({
            showValues: true,
            valueFormatter: valueWithPercentageFormatter
        })
        });
    }

    if (directorateChart) directorateChart.destroy();
    directorateChart = new Chart(document.getElementById("directorateChart"), {
        type: "bar",
        plugins: [ChartDataLabels, legendWhiteBackground],
        data: {
            labels: directorateRows.map(row => row.Diretoria),
            datasets: buildGroupedDatasets(directorateRows)
        },
        options: groupedChartOptions(15, true)
    });

    if (unitChart) unitChart.destroy();
    unitChart = new Chart(document.getElementById("unitChart"), {
        type: "bar",
        plugins: [ChartDataLabels, legendWhiteBackground],
        data: {
            labels: unitRows.map(row => row["Unidade/Agência"]),
            datasets: buildGroupedDatasets(unitRows)
        },
        options: groupedChartOptions()
    });
}


function exportIndicators(type) {
    window.location.href = `/export/indicators/${type}?${currentFilterParams().toString()}`;
}

function exportTable(type) {
    window.location.href = `/export/table/${type}?${currentFilterParams().toString()}`;
}

function exportSemaforo(type) {
    window.location.href = `/export/semaforo/${type}?${currentSemaforoFilterParams().toString()}`;
}

function exportSemaforoTable(type) {
    window.location.href = `/export/semaforo/table/${type}?${currentSemaforoFilterParams().toString()}`;
}

function renderSemaforoFilters(filters) {
    const filterBar = document.getElementById("semaforoFilterBar");
    if (!filterBar) return;

    const oldValues = {};
    document.querySelectorAll(".semaforo-filter-select").forEach(select => {
        oldValues[select.name] = select.value;
    });
    filterBar.innerHTML = "";

    semaforoFilterColumns.forEach(column => {
        const select = document.createElement("select");
        select.name = column;
        select.className = "filter-select semaforo-filter-select";
        select.innerHTML = `<option value="">${column}</option>`;

        (filters[column] || []).forEach(option => {
            const item = document.createElement("option");
            item.value = option;
            item.textContent = option;
            select.appendChild(item);
        });

        select.value = oldValues[column] || "";
        select.addEventListener("change", loadSemaforo);
        filterBar.appendChild(select);
    });
}

function clearSemaforoFilters() {
    document.querySelectorAll(".semaforo-filter-select").forEach(select => {
        select.value = "";
    });
    const search = document.getElementById("semaforoTableSearch");
    if (search) search.value = "";
    loadSemaforo();
}

function escapeHtml(value) {
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;")
        .replaceAll("\n", "<br>");
}

function formatSemaforoPercentage(value) {
    return Number(value || 0).toLocaleString("pt-BR", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function resetSemaforoCards() {
    const cards = [
        ["semaforoAttentionValue", "semaforoAttentionPercent"],
        ["semaforoMonitoringValue", "semaforoMonitoringPercent"],
        ["semaforoKnowledgeValue", "semaforoKnowledgePercent"]
    ];

    cards.forEach(([valueId, percentId]) => {
        const valueElement = document.getElementById(valueId);
        const percentElement = document.getElementById(percentId);
        if (valueElement) valueElement.textContent = "—";
        if (percentElement) percentElement.textContent = "% do Total";
    });
}

function renderSemaforoCards(indicators) {
    const conditions = indicators.condicoes || {};
    const cards = [
        {
            key: "Atenção",
            valueId: "semaforoAttentionValue",
            percentId: "semaforoAttentionPercent"
        },
        {
            key: "Monitoramento",
            valueId: "semaforoMonitoringValue",
            percentId: "semaforoMonitoringPercent"
        },
        {
            key: "Conhecimento",
            valueId: "semaforoKnowledgeValue",
            percentId: "semaforoKnowledgePercent"
        }
    ];

    cards.forEach(card => {
        const condition = conditions[card.key] || {};
        const quantity = Number(condition.quantidade || 0);
        const percentage = formatSemaforoPercentage(condition.percentual);
        const valueElement = document.getElementById(card.valueId);
        const percentElement = document.getElementById(card.percentId);

        if (valueElement) valueElement.textContent = quantity;
        if (percentElement) percentElement.textContent = `${percentage}% do Total`;
    });
}

function renderSemaforoRiskChart(indicators) {
    const canvas = document.getElementById("semaforoRiskChart");
    if (!canvas) return;

    const distribution = indicators.distribuicao_risco || {};
    const labels = ["Alto", "Significativo", "Moderado", "Baixo"];
    const values = labels.map(label => Number((distribution[label] || {}).quantidade || 0));
    const colors = ["#FF3333", "#F59A2F", "#F8FF25", "#76D84F"];

    if (semaforoRiskChart) semaforoRiskChart.destroy();

    semaforoRiskChart = new Chart(canvas, {
        type: "bar",
        plugins: [ChartDataLabels],
        data: {
            labels,
            datasets: [{
                data: values,
                backgroundColor: colors,
                borderColor: colors,
                borderWidth: 1,
                borderRadius: 8,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: "y",
            layout: { padding: { right: 24 } },
            plugins: {
                legend: { display: false },
                datalabels: {
                    display: true,
                    color: "#111111",
                    font: { weight: "800", size: 11 },
                    formatter: value => Number(value),
                    anchor: "end",
                    align: "right",
                    clamp: true
                },
                tooltip: {
                    callbacks: {
                        title: () => "",
                        label: context => `${labels[context.dataIndex]}: ${Number(context.raw || 0)}`
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0,
                        stepSize: 1,
                        color: "#111111"
                    },
                    grid: { color: "rgba(36, 54, 75, .12)" },
                    title: {
                        display: true,
                        text: "Quantidade de planos",
                        color: "#111111",
                        font: { weight: "700" }
                    }
                },
                y: {
                    grid: { display: false },
                    ticks: {
                        color: "#111111",
                        font: { weight: "800", size: 11 }
                    }
                }
            }
        }
    });
}

function wrapSemaforoText(value, maximum = 48) {
    const words = String(value || "").split(/\s+/).filter(Boolean);
    const lines = [];
    let line = "";
    words.forEach(word => {
        const candidate = line ? `${line} ${word}` : word;
        if (candidate.length > maximum && line) {
            lines.push(line);
            line = word;
        } else {
            line = candidate;
        }
    });
    if (line) lines.push(line);
    return lines;
}

function renderSemaforoRiskDistributionChart(indicators) {
    const canvas = document.getElementById("semaforoRiskDistributionChart");
    const placeholder = document.getElementById("semaforoRiskDistributionPlaceholder");
    if (!canvas || !placeholder) return;

    const distribution = indicators.distribuicao_risco || {};
    const labels = ["Alto", "Significativo", "Moderado", "Baixo"];
    const values = labels.map(label => Number((distribution[label] || {}).quantidade || 0));
    const colors = ["#FF3333", "#F59A2F", "#F8FF25", "#76D84F"];
    const total = values.reduce((sum, value) => sum + value, 0);

    if (semaforoRiskDistributionChart) {
        semaforoRiskDistributionChart.destroy();
        semaforoRiskDistributionChart = null;
    }

    if (!total) {
        canvas.hidden = true;
        placeholder.hidden = false;
        return;
    }

    placeholder.hidden = true;
    canvas.hidden = false;

    semaforoRiskDistributionChart = new Chart(canvas, {
        type: "pie",
        plugins: [ChartDataLabels, legendWhiteBackground],
        data: {
            labels,
            datasets: [{
                data: values,
                backgroundColor: colors,
                borderColor: "#FFFFFF",
                borderWidth: 3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true,
                    position: "bottom",
                    labels: {
                        color: "#111111",
                        usePointStyle: true,
                        padding: 10,
                        boxWidth: 10,
                        font: { weight: "700", size: 10 },
                        generateLabels: chart => labels.map((label, index) => ({
                            text: label,
                            fillStyle: colors[index],
                            strokeStyle: "#FFFFFF",
                            lineWidth: 1,
                            hidden: !chart.getDataVisibility(index),
                            index
                        }))
                    }
                },
                datalabels: {
                    color: "#111111",
                    font: { weight: "800", size: 11 },
                    formatter: value => {
                        if (!value || !total) return "";
                        return `${(value / total * 100).toLocaleString("pt-BR", {
                            minimumFractionDigits: 2,
                            maximumFractionDigits: 2
                        })}%`;
                    }
                },
                tooltip: {
                    callbacks: {
                        title: () => "",
                        label: context => {
                            const percentage = total ? Number(context.raw || 0) / total * 100 : 0;
                            return `${labels[context.dataIndex]}: ${percentage.toLocaleString("pt-BR", {
                                minimumFractionDigits: 2,
                                maximumFractionDigits: 2
                            })}%`;
                        }
                    }
                }
            }
        }
    });
}

function renderSemaforoConditionChart(indicators) {
    const container = document.getElementById("semaforoConditionChart");
    if (!container) return;
    const rules = indicators.regras || [];
    const categoryColors = {
        "Atenção": "#FF3333",
        "Monitoramento": "#FFDE59",
        "Conhecimento": "#76D84F"
    };
    container.innerHTML = "";
    rules.forEach(rule => {
        const quantity = Number(rule.quantidade || 0);
        const item = document.createElement("div");
        item.className = "condition-list-item";
        item.setAttribute("aria-label", `(${quantity}) ${rule.condicao}`);

        const dot = document.createElement("span");
        dot.className = "condition-value-dot";
        dot.textContent = `(${quantity})`;
        dot.style.backgroundColor = categoryColors[rule.categoria] || "#4F81BD";

        const label = document.createElement("span");
        label.className = "condition-list-label";
        label.textContent = rule.condicao;

        item.append(dot, label);
        container.appendChild(item);
    });
}

function renderSemaforoTable(rows) {
    const table = document.getElementById("semaforoDataTable");
    if (!table) return;
    table.querySelector("thead").innerHTML = `
        <tr>${semaforoTableColumns.map(column => `<th>${escapeHtml(column)}</th>`).join("")}</tr>
    `;
    table.querySelector("tbody").innerHTML = rows.map(row => `
        <tr>${semaforoTableColumns.map(column => `<td>${escapeHtml(row[column] ?? "")}</td>`).join("")}</tr>
    `).join("");
}

function clearSemaforoTableFilters() {
    clearSemaforoFilters();
}

function resetReincidenceCards() {
    [
        "reincidenceMonitoringValue",
        "reincidencePlansValue",
        "reincidenceRateValue"
    ].forEach(id => {
        const element = document.getElementById(id);
        if (element) element.textContent = "—";
    });
}

function renderReincidenceCards(indicators) {
    const recurrence = indicators.reincidencia || {};
    const monitoring = document.getElementById("reincidenceMonitoringValue");
    const recurrent = document.getElementById("reincidencePlansValue");
    const rate = document.getElementById("reincidenceRateValue");

    if (monitoring) monitoring.textContent = Number(recurrence.planos_monitoramento || 0);
    if (recurrent) recurrent.textContent = Number(recurrence.planos_reincidentes || 0);
    if (rate) rate.textContent = `${formatSemaforoPercentage(recurrence.taxa_reincidencia)}%`;
}

function renderReincidenceCriticalityChart(indicators) {
    const canvas = document.getElementById("reincidenceCriticalityChart");
    if (!canvas) return;

    const recurrence = indicators.reincidencia || {};
    const criticality = recurrence.criticidade_reincidente || {};
    const labels = ["Alto", "Significativo", "Moderado", "Baixo"];
    const values = labels.map(label => Number(criticality[label] || 0));
    const colors = ["#FF3333", "#F59A2F", "#FFDE59", "#76D84F"];

    if (recurrenceCriticalityChart) recurrenceCriticalityChart.destroy();

    recurrenceCriticalityChart = new Chart(canvas, {
        type: "bar",
        data: {
            labels,
            datasets: [{
                data: values,
                backgroundColor: colors,
                borderColor: colors,
                borderWidth: 1,
                borderRadius: 8,
                borderSkipped: false
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: "y",
            plugins: {
                legend: { display: false },
                datalabels: { display: false },
                tooltip: {
                    callbacks: {
                        title: () => "",
                        label: context => `${labels[context.dataIndex]}: ${context.raw}`
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: {
                        precision: 0,
                        stepSize: 1,
                        color: "#111111"
                    },
                    grid: { color: "rgba(36, 54, 75, .12)" },
                    title: {
                        display: true,
                        text: "Quantidade de planos",
                        color: "#111111",
                        font: { weight: "700" }
                    }
                },
                y: {
                    grid: { display: false },
                    ticks: {
                        color: "#111111",
                        font: { weight: "700" }
                    }
                }
            }
        }
    });
}

function renderReincidenceReportsChart(indicators) {
    const container = document.getElementById("reincidenceReportsChart");
    const placeholder = document.getElementById("reincidenceReportsPlaceholder");
    if (!container || !placeholder) return;

    const recurrence = indicators.reincidencia || {};
    const reports = recurrence.reincidentes_por_relatorio || [];
    const palette = [
        "#FDC01D",
        "#FF8A1A",
        "#F23B1D",
        "#A7265D"
    ];
    const total = reports.reduce((sum, item) => sum + Number(item.quantidade || 0), 0);
    container.innerHTML = "";

    if (!total) {
        container.hidden = true;
        placeholder.hidden = false;
        return;
    }

    placeholder.hidden = true;
    container.hidden = false;
    reports.forEach((report, index) => {
        const quantity = Number(report.quantidade || 0);
        const name = String(report.relatorio || "");
        const item = document.createElement("div");
        item.className = "condition-list-item";
        item.setAttribute("aria-label", `(${quantity}) ${name}`);

        const dot = document.createElement("span");
        dot.className = "condition-value-dot";
        dot.textContent = `(${quantity})`;
        dot.style.backgroundColor = palette[index % palette.length];

        const label = document.createElement("span");
        label.className = "condition-list-label";
        label.textContent = name;

        item.append(dot, label);
        container.appendChild(item);
    });
}

async function loadSemaforo() {
    if (!document.getElementById("semaforoDashboard")) return;

    const message = document.getElementById("semaforoMessage");

    try {
        const response = await fetch(`/api/semaforo?${currentSemaforoFilterParams().toString()}`);
        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.error || "Não foi possível carregar o semáforo.");
        }

        if (message) {
            message.hidden = true;
            message.textContent = "";
        }

        semaforoRows = data.rows || [];
        semaforoTableColumns = data.table_columns || [];
        renderSemaforoFilters(data.filters || {});
        const activePlansLabel = document.getElementById("semaforoActivePlansLabel");
        if (activePlansLabel) {
            const activeFilters = Object.entries(data.active_filters || {});
            const filterLabel = activeFilters.length
                ? ` | ${activeFilters.map(([key, value]) => `${key}: ${value}`).join(" | ")}`
                : "";
            activePlansLabel.textContent = `Tabela Analítica - ${semaforoRows.length} Planos Ativos${filterLabel}`;
        }

        renderSemaforoCards(data.indicators);
        renderSemaforoRiskChart(data.indicators);
        renderSemaforoRiskDistributionChart(data.indicators);
        renderSemaforoConditionChart(data.indicators);
        renderReincidenceCards(data.indicators);
        renderReincidenceCriticalityChart(data.indicators);
        renderReincidenceReportsChart(data.indicators);
        renderSemaforoTable(semaforoRows);
    } catch (error) {
        console.error(error);
        resetSemaforoCards();
        resetReincidenceCards();
        semaforoRows = [];
        semaforoTableColumns = [];
        renderSemaforoTable([]);

        if (message) {
            message.textContent = error.message;
            message.hidden = false;
        }
    }
}

function renderPaintFilters(filters) {
    const filterBar = document.getElementById("paintFilterBar");
    if (!filterBar) return;

    const oldValues = {};
    document.querySelectorAll(".paint-filter-select").forEach(select => {
        oldValues[select.name] = select.value;
    });
    filterBar.innerHTML = "";

    paintFilterColumns.forEach(column => {
        const select = document.createElement("select");
        select.name = column;
        select.className = "filter-select paint-filter-select";
        select.innerHTML = `<option value="">${escapeHtml(paintFilterLabels[column] || column)}</option>`;

        (filters[column] || []).forEach(option => {
            const item = document.createElement("option");
            item.value = option;
            item.textContent = option;
            select.appendChild(item);
        });

        select.value = oldValues[column] || "";
        select.addEventListener("change", loadPaint);
        filterBar.appendChild(select);
    });
}

function clearPaintFilters() {
    document.querySelectorAll(".paint-filter-select").forEach(select => {
        select.value = "";
    });
    const search = document.getElementById("paintTableSearch");
    if (search) search.value = "";
    loadPaint();
}

function clearPaintTableFilters() {
    clearPaintFilters();
}

function exportPaint(type) {
    window.location.href = `/export/paint/${type}?${currentPaintFilterParams().toString()}`;
}

function exportPaintTable(type) {
    window.location.href = `/export/paint/table/${type}?${currentPaintFilterParams().toString()}`;
}

function renderPaintCards(indicators) {
    const values = {
        paintTotal: indicators.total_auditorias,
        paintCompleted: indicators.concluidas,
        paintRunning: indicators.em_andamento,
        paintNotStarted: indicators.nao_iniciadas
    };
    Object.entries(values).forEach(([id, value]) => {
        const element = document.getElementById(id);
        if (element) element.textContent = Number(value || 0);
    });
}

function renderPaintProcessStatusChart(indicators) {
    const canvas = document.getElementById("paintProcessStatusChart");
    if (!canvas) return;

    const rows = indicators.processos_status || [];
    const statuses = [
        { label: "Concluída", color: "#76D84F" },
        { label: "Em Andamento", color: "#FFDE59" },
        { label: "Não Iniciada", color: "#FF3333" }
    ];
    if (paintProcessStatusChart) paintProcessStatusChart.destroy();

    paintProcessStatusChart = new Chart(canvas, {
        type: "bar",
        plugins: [ChartDataLabels, legendWhiteBackground],
        data: {
            labels: rows.map(row => row.processo),
            datasets: statuses.map(status => ({
                label: status.label,
                data: rows.map(row => row.status === status.label ? 0.66 : 0),
                backgroundColor: status.color,
                borderColor: "#FFFFFF",
                borderWidth: 1,
                borderRadius: 8,
                borderSkipped: false,
                statusLabel: status.label
            }))
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: "y",
            layout: { padding: { top: 8, right: 12, bottom: 18, left: 8 } },
            plugins: {
                legend: {
                    position: "bottom",
                    labels: {
                        color: "#111111",
                        usePointStyle: true,
                        padding: 14,
                        boxWidth: 10,
                        font: { size: 10, weight: "700" }
                    }
                },
                datalabels: {
                    color: "#111111",
                    font: { size: 9, weight: "800" },
                    formatter: (value, context) => Number(value) > 0
                        ? context.dataset.statusLabel
                        : "",
                    anchor: "center",
                    align: "center",
                    clamp: true
                },
                tooltip: {
                    callbacks: {
                        title: items => items.length ? rows[items[0].dataIndex]?.processo || "" : "",
                        label: context => Number(context.raw || 0) > 0 ? context.dataset.label : ""
                    },
                    filter: context => Number(context.raw || 0) > 0
                }
            },
            scales: {
                x: {
                    stacked: true,
                    min: 0,
                    max: 1,
                    display: false,
                    grid: { display: false }
                },
                y: {
                    stacked: true,
                    grid: { display: false },
                    afterFit: scale => {
                        const maximumWidth = Math.floor(scale.chart.width * 0.56);
                        scale.width = Math.min(maximumWidth, Math.max(scale.width, 380));
                    },
                    ticks: {
                        color: "#111111",
                        autoSkip: false,
                        padding: 6,
                        font: { size: 12, weight: "700" }
                    },
                    border: { display: false }
                }
            }
        }
    });
}

function renderPaintStatusDistributionChart(indicators) {
    const canvas = document.getElementById("paintStatusDistributionChart");
    if (!canvas) return;

    const labels = ["Concluída", "Em Andamento", "Não Iniciada"];
    const colors = ["#76D84F", "#FFDE59", "#FF3333"];
    const distribution = indicators.distribuicao_status || {};
    const values = labels.map(label => Number(distribution[label]?.quantidade || 0));

    if (paintStatusDistributionChart) paintStatusDistributionChart.destroy();
    paintStatusDistributionChart = new Chart(canvas, {
        type: "pie",
        plugins: [ChartDataLabels, legendWhiteBackground],
        data: {
            labels,
            datasets: [{
                data: values,
                backgroundColor: colors,
                borderColor: "#FFFFFF",
                borderWidth: 3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: { top: 8, right: 8, bottom: 18, left: 8 } },
            plugins: {
                legend: {
                    position: "bottom",
                    labels: {
                        color: "#111111",
                        usePointStyle: true,
                        padding: 14,
                        boxWidth: 10,
                        font: { size: 10, weight: "700" }
                    }
                },
                datalabels: {
                    color: "#111111",
                    font: { size: 12, weight: "800" },
                    formatter: value => {
                        const total = values.reduce((sum, item) => sum + item, 0);
                        if (!value || !total) return "";
                        return `${(Number(value) / total * 100).toLocaleString("pt-BR", {
                            minimumFractionDigits: 1,
                            maximumFractionDigits: 1
                        })}%`;
                    }
                },
                tooltip: {
                    callbacks: {
                        label: context => {
                            const total = values.reduce((sum, item) => sum + item, 0);
                            const percent = total ? Number(context.raw || 0) / total * 100 : 0;
                            return `${context.label}: ${context.raw} (${percent.toLocaleString("pt-BR", {
                                minimumFractionDigits: 1,
                                maximumFractionDigits: 1
                            })}%)`;
                        }
                    }
                }
            }
        }
    });
}

function renderPaintTable(rows) {
    const table = document.getElementById("paintTable");
    if (!table) return;

    table.querySelector("thead").innerHTML = `
        <tr>${paintTableColumns.map(column => `<th>${escapeHtml(column)}</th>`).join("")}</tr>
    `;
    table.querySelector("tbody").innerHTML = rows.map(row => `
        <tr>${paintTableColumns.map(column => `<td>${escapeHtml(row[column] ?? "")}</td>`).join("")}</tr>
    `).join("");
}

async function loadPaint() {
    if (!document.getElementById("paintDashboard")) return;
    const message = document.getElementById("paintMessage");

    try {
        const response = await fetch(`/api/paint?${currentPaintFilterParams().toString()}`);
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "Não foi possível carregar a execução do PAINT.");

        paintRows = data.rows || [];
        paintTableColumns = data.table_columns || [];
        renderPaintFilters(data.filters || {});
        renderPaintCards(data.indicators || {});
        renderPaintProcessStatusChart(data.indicators || {});
        renderPaintStatusDistributionChart(data.indicators || {});
        renderPaintTable(paintRows);

        const label = document.getElementById("paintTableLabel");
        if (label) {
            const activeFilters = Object.entries(data.active_filters || {});
            const filterText = activeFilters.length
                ? ` | ${activeFilters.map(([key, value]) => `${paintFilterLabels[key] || key}: ${value}`).join(" | ")}`
                : "";
            label.textContent = `Tabela Analítica - ${paintRows.length} Auditorias${filterText}`;
        }
        if (message) {
            message.hidden = true;
            message.textContent = "";
        }
    } catch (error) {
        console.error(error);
        paintRows = [];
        paintTableColumns = [];
        renderPaintTable([]);
        if (message) {
            message.hidden = false;
            message.textContent = error.message;
        }
    }
}

function renderCarteiraFilters(filters) {
    const filterBar = document.getElementById("carteiraFilterBar");
    if (!filterBar) return;

    const oldValues = {};
    document.querySelectorAll(".carteira-filter-select").forEach(select => {
        oldValues[select.name] = select.value;
    });
    filterBar.innerHTML = "";

    carteiraFilterColumns.forEach(column => {
        const select = document.createElement("select");
        select.name = column;
        select.className = "filter-select carteira-filter-select";
        select.innerHTML = `<option value="">${escapeHtml(column)}</option>`;

        (filters[column] || []).forEach(option => {
            const item = document.createElement("option");
            item.value = option;
            item.textContent = option;
            select.appendChild(item);
        });

        select.value = oldValues[column] || "";
        select.addEventListener("change", loadCarteira);
        filterBar.appendChild(select);
    });
}

function clearCarteiraFilters() {
    document.querySelectorAll(".carteira-filter-select").forEach(select => {
        select.value = "";
    });
    const search = document.getElementById("carteiraTableSearch");
    if (search) search.value = "";
    loadCarteira();
}

function clearCarteiraTableFilters() {
    clearCarteiraFilters();
}

function exportCarteira(type) {
    window.location.href = `/export/carteira/${type}?${currentCarteiraFilterParams().toString()}`;
}

function exportCarteiraTable(type) {
    const params = currentCarteiraFilterParams();
    params.set("table", carteiraActiveTable);
    window.location.href = `/export/carteira/table/${type}?${params.toString()}`;
}

function renderCarteiraCards(indicators) {
    const values = {
        carteiraTotal: indicators.total_auditorias,
        carteiraMonitoring: indicators.planos_monitoramento,
        carteiraClosed: indicators.encerradas,
        carteiraAwaiting: indicators.aguardando_plano
    };
    Object.entries(values).forEach(([id, value]) => {
        const element = document.getElementById(id);
        if (element) element.textContent = Number(value || 0);
    });
}

function renderCarteiraStatusChart(indicators) {
    const canvas = document.getElementById("carteiraStatusChart");
    if (!canvas) return;

    const labels = [
        "Auditorias com Planos em Monitoramento",
        "Auditorias Encerradas",
        "Aguardando Plano de Ação"
    ];
    const colors = ["#5B9BD5", "#76D84F", "#FF3333"];
    const distribution = indicators.distribuicao_status || {};
    const values = labels.map(label => Number(distribution[label]?.quantidade || 0));

    if (carteiraStatusChart) carteiraStatusChart.destroy();
    carteiraStatusChart = new Chart(canvas, {
        type: "pie",
        plugins: [ChartDataLabels, legendWhiteBackground],
        data: {
            labels,
            datasets: [{
                data: values,
                backgroundColor: colors,
                borderColor: "#FFFFFF",
                borderWidth: 3
            }]
        },
        options: carteiraPieOptions(values)
    });
}

function renderCarteiraPaintChart(indicators) {
    const canvas = document.getElementById("carteiraPaintChart");
    if (!canvas) return;

    const distribution = indicators.distribuicao_paint || {};
    const labels = Object.keys(distribution);
    const values = labels.map(label => Number(distribution[label]?.quantidade || 0));
    const palette = ["#6C63E8", "#FF5C7A", "#27C2A4", "#FF9F1C", "#2E86DE", "#C846D9"];

    if (carteiraPaintChart) carteiraPaintChart.destroy();
    carteiraPaintChart = new Chart(canvas, {
        type: "pie",
        plugins: [ChartDataLabels, legendWhiteBackground],
        data: {
            labels,
            datasets: [{
                data: values,
                backgroundColor: labels.map((_, index) => palette[index % palette.length]),
                borderColor: "#FFFFFF",
                borderWidth: 3
            }]
        },
        options: carteiraPieOptions(values)
    });
}

function carteiraPieOptions(values) {
    return {
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: { top: 8, right: 8, bottom: 22, left: 8 } },
        plugins: {
            legend: {
                position: "bottom",
                labels: {
                    color: "#111111",
                    usePointStyle: true,
                    padding: 14,
                    boxWidth: 10,
                    font: { size: 10, weight: "700" }
                }
            },
            datalabels: {
                color: "#111111",
                font: { size: 12, weight: "800" },
                formatter: value => {
                    const total = values.reduce((sum, item) => sum + item, 0);
                    if (!value || !total) return "";
                    return `${(Number(value) / total * 100).toLocaleString("pt-BR", {
                        minimumFractionDigits: 1,
                        maximumFractionDigits: 1
                    })}%`;
                }
            },
            tooltip: {
                callbacks: {
                    label: context => {
                        const total = values.reduce((sum, item) => sum + item, 0);
                        const percent = total ? Number(context.raw || 0) / total * 100 : 0;
                        return `${context.label}: ${context.raw} (${percent.toLocaleString("pt-BR", {
                            minimumFractionDigits: 1,
                            maximumFractionDigits: 1
                        })}%)`;
                    }
                }
            }
        }
    };
}

function renderCarteiraTable(rows = null) {
    const table = document.getElementById("carteiraTable");
    if (!table) return;

    const selectedTable = carteiraTables[carteiraActiveTable] || { columns: [], rows: [] };
    const columns = selectedTable.columns || [];
    const visibleRows = rows || selectedTable.rows || [];
    table.querySelector("thead").innerHTML = `
        <tr>${columns.map(column => `<th>${escapeHtml(column)}</th>`).join("")}</tr>
    `;
    table.querySelector("tbody").innerHTML = visibleRows.map(row => `
        <tr>${columns.map(column => {
            const normalizedStatus = String(row[column] || "")
                .normalize("NFD")
                .replace(/[\u0300-\u036f]/g, "")
                .toLowerCase()
                .replaceAll(" ", "-");
            const statusClass = carteiraActiveTable === "aguardando" && column === "Status" && normalizedStatus === "vencido"
                ? "carteira-status-vencido"
                : "";
            return `<td class="${statusClass}">${escapeHtml(row[column] ?? "")}</td>`;
        }).join("")}</tr>
    `).join("");

    document.querySelectorAll(".carteira-folder-tab").forEach(tab => {
        const active = tab.dataset.table === carteiraActiveTable;
        tab.classList.toggle("active", active);
        tab.setAttribute("aria-selected", active ? "true" : "false");
    });
    const folder = document.querySelector(".carteira-folder-table");
    if (folder) {
        folder.classList.toggle("carteira-folder-main", carteiraActiveTable === "carteira");
        folder.classList.toggle("carteira-folder-awaiting", carteiraActiveTable === "aguardando");
    }
}

function selectCarteiraTable(tableKey) {
    carteiraActiveTable = tableKey === "aguardando" ? "aguardando" : "carteira";
    const search = document.getElementById("carteiraTableSearch");
    if (search) search.value = "";
    renderCarteiraTable();
}

async function loadCarteira() {
    if (!document.getElementById("carteiraDashboard")) return;
    const message = document.getElementById("carteiraMessage");

    try {
        const response = await fetch(`/api/carteira?${currentCarteiraFilterParams().toString()}`);
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || "Não foi possível carregar a Carteira de Auditoria.");

        carteiraTables = data.tables || {
            carteira: { columns: [], rows: [] },
            aguardando: { columns: [], rows: [] }
        };
        renderCarteiraFilters(data.filters || {});
        renderCarteiraCards(data.indicators || {});
        renderCarteiraStatusChart(data.indicators || {});
        renderCarteiraPaintChart(data.indicators || {});
        renderCarteiraTable();

        const label = document.getElementById("carteiraTableLabel");
        if (label) {
            const total = carteiraTables.carteira?.rows?.length || 0;
            label.textContent = `Tabelas Analíticas - Total da Carteira (${total})`;
        }
        if (message) {
            message.hidden = true;
            message.textContent = "";
        }
    } catch (error) {
        console.error(error);
        carteiraTables = {
            carteira: { columns: [], rows: [] },
            aguardando: { columns: [], rows: [] }
        };
        renderCarteiraTable();
        if (message) {
            message.hidden = false;
            message.textContent = error.message;
        }
    }
}

document.addEventListener("click", event => {
    const tab = event.target.closest(".carteira-folder-tab");
    if (tab) selectCarteiraTable(tab.dataset.table);
});

document.addEventListener("input", event => {
    const value = event.target.value.toLowerCase();
    if (event.target.id === "tableSearch") {
        const filtered = dashboardRows.filter(row =>
            Object.values(row).some(cell => String(cell).toLowerCase().includes(value))
        );
        renderTable(filtered);
    }
    if (event.target.id === "semaforoTableSearch") {
        const filtered = semaforoRows.filter(row =>
            Object.values(row).some(cell => String(cell).toLowerCase().includes(value))
        );
        renderSemaforoTable(filtered);
    }
    if (event.target.id === "paintTableSearch") {
        const filtered = paintRows.filter(row =>
            Object.values(row).some(cell => String(cell).toLowerCase().includes(value))
        );
        renderPaintTable(filtered);
    }
    if (event.target.id === "carteiraTableSearch") {
        const selectedTable = carteiraTables[carteiraActiveTable] || { rows: [] };
        const filtered = selectedTable.rows.filter(row =>
            Object.values(row).some(cell => String(cell).toLowerCase().includes(value))
        );
        renderCarteiraTable(filtered);
    }
});

document.addEventListener("DOMContentLoaded", loadDashboard);
document.addEventListener("DOMContentLoaded", loadSemaforo);
document.addEventListener("DOMContentLoaded", loadPaint);
document.addEventListener("DOMContentLoaded", loadCarteira);
