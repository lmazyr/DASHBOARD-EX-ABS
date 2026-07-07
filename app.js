document.addEventListener("DOMContentLoaded", () => {
    const COLORS = {
        primary: "#4f6f52",
        success: "#16a34a",
        warning: "#f59e0b",
        danger: "#dc2626",
        info: "#2563eb",
        neutral: "#64748b"
    };

    const DEFAULT_WORKBOOK_URL = "./Absenteísmo -Expedição Junho 2026.xlsx";
    const DEFAULT_ONEDRIVE_LINK = "https://scaniaazureservices-my.sharepoint.com/:x:/r/personal/lucas_martins_scania_com/Documents/DASHBORD%20%20AB/Absente%C3%ADsmo%20-Expedi%C3%A7%C3%A3o%20Junho%202026.xlsx?d=w3a62a05b344044cd951d2eb4cd9d4280&csf=1&web=1&e=8kEhEM";
    const CACHE_KEY = "seseDashboardCache";
    const LINK_KEY = "oneDriveLink";
    const LOCAL_HOSTS = ["localhost", "127.0.0.1", ""];
    const MONTH_ORDER = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"];
    const MONTH_LABELS = {
        JANEIRO: "Janeiro",
        FEVEREIRO: "Fevereiro",
        "MARÇO": "Março",
        ABRIL: "Abril",
        MAIO: "Maio",
        JUNHO: "Junho",
        JULHO: "Julho",
        AGOSTO: "Agosto",
        SETEMBRO: "Setembro",
        OUTUBRO: "Outubro",
        NOVEMBRO: "Novembro",
        DEZEMBRO: "Dezembro"
    };

    const state = {
        workbookMonths: {},
        dadosOriginais: [],
        dadosFiltrados: [],
        activeMonth: "",
        activeSheet: "",
        sourceLabel: "aguardando dados",
        autoRefreshInterval: null,
        isAutoRefreshActive: false,
        refreshIntervalTime: 30000,
        charts: {
            linha: null,
            pizza: null,
            ranking: null
        }
    };

    const $ = (selector) => document.querySelector(selector);
    const elements = {
        searchInput: $("#searchInput"),
        turnoFilter: $("#turnoFilter"),
        funcaoFilter: $("#funcaoFilter"),
        btnFiltro: $("#btnFiltro"),
        btnUploadExcel: $("#btnUploadExcel"),
        excelFile: $("#excelFile"),
        btnExportarExcel: $("#btnExportarExcel"),
        btnExportPdf: $("#btnExportPdf"),
        btnAutoRefresh: $("#btnAutoRefresh"),
        refreshInterval: $("#refreshInterval"),
        btnSalvarLink: $("#btnSalvarLink"),
        oneDriveLink: $("#oneDriveLink"),
        sidebar: $(".sidebar"),
        toggleSidebar: $(".toggle-sidebar"),
        main: $(".main"),
        themeBtn: $(".theme-btn"),
        mesPrincipal: $("#mesPrincipal"),
        mesComparacao: $("#mesComparacao"),
        syncStatus: $("#syncStatus"),
        dataSourceLabel: $("#dataSourceLabel"),
        lastUpdateLabel: $("#lastUpdateLabel"),
        activeSheetLabel: $("#activeSheetLabel"),
        comparativoMeses: $("#comparativoMeses"),
        tableBody: $("#tableBody")
    };

    function normalizeText(value) {
        return String(value ?? "")
            .trim()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, "")
            .toUpperCase();
    }

    function toNumber(value) {
        if (typeof value === "number") return value;
        const parsed = Number(String(value ?? "0").replace(",", ".").trim());
        return Number.isFinite(parsed) ? parsed : 0;
    }

    function escapeHtml(value) {
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    function formatTime(date = new Date()) {
        return date.toLocaleString("pt-BR", {
            day: "2-digit",
            month: "2-digit",
            year: "numeric",
            hour: "2-digit",
            minute: "2-digit",
            second: "2-digit"
        });
    }

    function setSyncStatus(message, type = "neutral") {
        if (!elements.syncStatus) return;
        const colorByType = {
            success: COLORS.success,
            error: COLORS.danger,
            warning: COLORS.warning,
            neutral: COLORS.neutral
        };
        elements.syncStatus.textContent = message;
        elements.syncStatus.style.color = colorByType[type] || COLORS.neutral;
    }

    function setLoading(isLoading) {
        const container = $(".refresh-container");
        container?.classList.toggle("syncing", isLoading);
    }

    function getMetric(item, key) {
        return toNumber(item[key]);
    }

    function getTotalAusencias(item) {
        return getMetric(item, "FN") + getMetric(item, "FJ") + getMetric(item, "A");
    }

    function getTotalRegistros(item) {
        return getMetric(item, "P") + getMetric(item, "BH") + getMetric(item, "FN") + getMetric(item, "FJ") + getMetric(item, "A") + getMetric(item, "T") + getMetric(item, "S") + getMetric(item, "F");
    }

    function convertOneDriveLink(link) {
        if (!link) return "";
        const clean = link.trim();
        if (!clean.includes("sharepoint.com") && !clean.includes("onedrive.live.com")) return clean;
        if (clean.includes("download=1")) return clean;

        const separator = clean.includes("?") ? "&" : "?";
        return `${clean}${separator}download=1`;
    }

    function isLocalDashboard() {
        return LOCAL_HOSTS.includes(window.location.hostname);
    }

    async function fetchWorkbook(url) {
        const cacheBuster = `v=${Date.now()}`;
        const separator = url.includes("?") ? "&" : "?";
        const response = await fetch(`${url}${separator}${cacheBuster}`, { cache: "no-store" });
        if (!response.ok) throw new Error(`Falha ao carregar Excel (${response.status})`);
        return response.arrayBuffer();
    }

    function readWorkbookFromBuffer(buffer) {
        return XLSX.read(new Uint8Array(buffer), {
            type: "array",
            cellDates: true,
            raw: false
        });
    }

    function findHeaderRow(rows) {
        return rows.findIndex((row) => row.some((cell) => normalizeText(cell) === "TURNO"));
    }

    function normalizeRow(headers, row) {
        const normalized = {};

        headers.forEach((header, index) => {
            const key = normalizeText(header);
            const value = row[index] ?? "";

            if (key === "TURNO") normalized.TURNO = String(value).trim();
            if (key === "NOME COMPLETO" || key === "NOME") normalized["NOME COMPLETO"] = String(value).trim();
            if (key === "FUNCAO" || key === "FUNÇÃO") normalized["FUNÇÃO"] = String(value).trim();
            if (["P", "BH", "FN", "FJ", "A", "T", "S", "F"].includes(key)) normalized[key] = toNumber(value);
        });

        normalized.TURNO = normalized.TURNO || "";
        normalized["NOME COMPLETO"] = normalized["NOME COMPLETO"] || "";
        normalized["FUNÇÃO"] = normalized["FUNÇÃO"] || "";

        return normalized;
    }

    function parseMonthSheet(workbook, sheetName) {
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        const headerIndex = findHeaderRow(rows);
        if (headerIndex < 0) return [];

        const headers = rows[headerIndex];
        return rows
            .slice(headerIndex + 1)
            .map((row) => normalizeRow(headers, row))
            .filter((item) => item.TURNO && item["NOME COMPLETO"]);
    }

    function parseWorkbook(workbook) {
        const months = {};

        workbook.SheetNames.forEach((sheetName) => {
            const normalizedName = normalizeText(sheetName);
            const month = MONTH_ORDER.find((name) => normalizedName.includes(normalizeText(name)));
            if (!month || normalizedName === "RESUMO") return;

            const rows = parseMonthSheet(workbook, sheetName);
            if (rows.length) {
                months[month] = {
                    sheetName,
                    rows
                };
            }
        });

        if (!Object.keys(months).length) {
            throw new Error("Nenhuma aba mensal com cabeçalho TURNO foi encontrada.");
        }

        return months;
    }

    function chooseDefaultMonth(months) {
        const available = Object.keys(months);
        return available.sort((a, b) => MONTH_ORDER.indexOf(b) - MONTH_ORDER.indexOf(a))[0];
    }

    function populateMonthSelects() {
        const available = Object.keys(state.workbookMonths)
            .sort((a, b) => MONTH_ORDER.indexOf(a) - MONTH_ORDER.indexOf(b));

        elements.mesPrincipal.innerHTML = available
            .map((month) => `<option value="${month}">${MONTH_LABELS[month] || month}</option>`)
            .join("");

        elements.mesComparacao.innerHTML = '<option value="">Sem comparação</option>' + available
            .map((month) => `<option value="${month}">${MONTH_LABELS[month] || month}</option>`)
            .join("");

        if (state.activeMonth && available.includes(state.activeMonth)) {
            elements.mesPrincipal.value = state.activeMonth;
        }
    }

    function populateFilters() {
        const turnos = [...new Set(state.dadosOriginais.map((item) => item.TURNO).filter(Boolean))].sort();
        const funcoes = [...new Set(state.dadosOriginais.map((item) => item["FUNÇÃO"]).filter(Boolean))].sort();

        const currentTurno = elements.turnoFilter.value;
        const currentFuncao = elements.funcaoFilter.value;

        elements.turnoFilter.innerHTML = '<option value="">Todos</option>' + turnos
            .map((turno) => `<option value="${escapeHtml(turno)}">${escapeHtml(turno)}</option>`)
            .join("");

        elements.funcaoFilter.innerHTML = '<option value="">Todas</option>' + funcoes
            .map((funcao) => `<option value="${escapeHtml(funcao)}">${escapeHtml(funcao)}</option>`)
            .join("");

        if (turnos.includes(currentTurno)) elements.turnoFilter.value = currentTurno;
        if (funcoes.includes(currentFuncao)) elements.funcaoFilter.value = currentFuncao;
    }

    function setActiveMonth(month) {
        const selected = month || chooseDefaultMonth(state.workbookMonths);
        const monthData = state.workbookMonths[selected];
        if (!monthData) return;

        state.activeMonth = selected;
        state.activeSheet = monthData.sheetName;
        state.dadosOriginais = [...monthData.rows];

        if (elements.mesPrincipal.value !== selected) {
            elements.mesPrincipal.value = selected;
        }

        populateFilters();
        applyFilters();
    }

    function loadWorkbookData(workbook, sourceLabel) {
        const previousMonth = state.activeMonth;
        state.workbookMonths = parseWorkbook(workbook);
        state.sourceLabel = sourceLabel;

        populateMonthSelects();
        setActiveMonth(previousMonth && state.workbookMonths[previousMonth] ? previousMonth : chooseDefaultMonth(state.workbookMonths));
        saveCache();
        updateStatusLabels();
    }

    function saveCache() {
        const payload = {
            workbookMonths: state.workbookMonths,
            activeMonth: state.activeMonth,
            sourceLabel: state.sourceLabel,
            savedAt: new Date().toISOString()
        };
        localStorage.setItem(CACHE_KEY, JSON.stringify(payload));
    }

    function loadCache() {
        try {
            const raw = localStorage.getItem(CACHE_KEY);
            if (!raw) return false;

            const cached = JSON.parse(raw);
            if (!cached.workbookMonths || !Object.keys(cached.workbookMonths).length) return false;

            state.workbookMonths = cached.workbookMonths;
            state.activeMonth = cached.activeMonth;
            state.sourceLabel = `${cached.sourceLabel || "cache"} (cache)`;
            populateMonthSelects();
            setActiveMonth(state.activeMonth);
            updateStatusLabels(cached.savedAt ? new Date(cached.savedAt) : new Date());
            setSyncStatus("Cache", "warning");
            return true;
        } catch (error) {
            console.warn("Cache inválido:", error);
            return false;
        }
    }

    async function loadFromLocalWorkbook() {
        const buffer = await fetchWorkbook(DEFAULT_WORKBOOK_URL);
        loadWorkbookData(readWorkbookFromBuffer(buffer), "Excel da pasta");
        setSyncStatus("Atualizado", "success");
    }

    async function loadFromOneDrive() {
        const link = localStorage.getItem(LINK_KEY) || DEFAULT_ONEDRIVE_LINK;
        if (!link) throw new Error("Nenhum link do OneDrive configurado.");

        const buffer = await fetchWorkbook(convertOneDriveLink(link));
        loadWorkbookData(readWorkbookFromBuffer(buffer), "OneDrive");
        setSyncStatus("Atualizado", "success");
    }

    async function refreshData({ silent = false } = {}) {
        setLoading(true);
        if (!silent) setSyncStatus("Sincronizando...", "neutral");

        const sources = isLocalDashboard()
            ? [loadFromLocalWorkbook, loadFromOneDrive]
            : [loadFromOneDrive, loadFromLocalWorkbook];

        const errors = [];

        try {
            for (const loadSource of sources) {
                try {
                    await loadSource();
                    return;
                } catch (error) {
                    errors.push(error);
                }
            }

            if (!loadCache()) {
                renderEmptyState("Importe um Excel para iniciar o dashboard.");
            }

            setSyncStatus("Offline", "error");
            console.warn("Não foi possível atualizar:", errors);
        } finally {
            setLoading(false);
        }
    }

    function readLocalFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (event) => resolve(event.target.result);
            reader.onerror = () => reject(new Error("Falha ao ler o arquivo selecionado."));
            reader.readAsArrayBuffer(file);
        });
    }

    async function handleExcelUpload(event) {
        const file = event.target.files?.[0];
        if (!file) return;

        setLoading(true);
        setSyncStatus("Importando...", "neutral");

        try {
            const buffer = await readLocalFile(file);
            loadWorkbookData(readWorkbookFromBuffer(buffer), `Arquivo importado: ${file.name}`);
            setSyncStatus("Importado", "success");
        } catch (error) {
            alert(`Não foi possível importar o Excel: ${error.message}`);
            setSyncStatus("Erro", "error");
        } finally {
            setLoading(false);
            event.target.value = "";
        }
    }

    function applyFilters() {
        const nome = normalizeText(elements.searchInput.value);
        const turno = elements.turnoFilter.value;
        const funcao = elements.funcaoFilter.value;

        state.dadosFiltrados = state.dadosOriginais.filter((item) => {
            const nomeOk = !nome || normalizeText(item["NOME COMPLETO"]).includes(nome);
            const turnoOk = !turno || item.TURNO === turno;
            const funcaoOk = !funcao || item["FUNÇÃO"] === funcao;
            return nomeOk && turnoOk && funcaoOk;
        });

        updateDashboard();
    }

    function updateDashboard() {
        updateKPIs();
        updateTable();
        updateLineChart();
        updateDonutChart();
        updateRankingChart();
        updateComparison();
        updateStatusLabels();
    }

    function updateStatusLabels(date = new Date()) {
        elements.dataSourceLabel.textContent = `Fonte: ${state.sourceLabel}`;
        elements.lastUpdateLabel.textContent = `Última atualização: ${formatTime(date)}`;
        elements.activeSheetLabel.textContent = `Aba: ${state.activeSheet || "--"}`;
    }

    function updateKPIs() {
        const totalFuncionarios = state.dadosFiltrados.length;
        const totalFJ = state.dadosFiltrados.reduce((sum, item) => sum + getMetric(item, "FJ"), 0);
        const totalFN = state.dadosFiltrados.reduce((sum, item) => sum + getMetric(item, "FN"), 0);
        const totalAusencias = state.dadosFiltrados.reduce((sum, item) => sum + getTotalAusencias(item), 0);
        const totalPossivel = state.dadosFiltrados.reduce((sum, item) => sum + getTotalRegistros(item), 0);
        const presenca = totalPossivel > 0 ? ((totalPossivel - totalAusencias) / totalPossivel) * 100 : 0;

        $("#kpiFuncionarios").textContent = totalFuncionarios;
        $("#kpiFJ").textContent = totalFJ;
        $("#kpiFN").textContent = totalFN;
        $("#kpiPresenca").textContent = `${Math.max(0, presenca).toFixed(1)}%`;
    }

    function updateTable() {
        if (!state.dadosFiltrados.length) {
            renderEmptyState("Nenhum registro encontrado para os filtros atuais.");
            return;
        }

        elements.tableBody.innerHTML = state.dadosFiltrados.map((item) => {
            const total = getTotalAusencias(item);
            return `
                <tr>
                    <td>${escapeHtml(item["NOME COMPLETO"])}</td>
                    <td>${escapeHtml(item["FUNÇÃO"])}</td>
                    <td>${escapeHtml(item.TURNO)}</td>
                    <td><span class="badge badge-success">${getMetric(item, "P")}</span></td>
                    <td><span class="badge badge-neutral">${getMetric(item, "BH")}</span></td>
                    <td><span class="badge badge-warning">${getMetric(item, "FJ")}</span></td>
                    <td><span class="badge badge-danger">${getMetric(item, "FN")}</span></td>
                    <td><span class="badge badge-info">${getMetric(item, "A")}</span></td>
                    <td><span class="badge badge-neutral">${total}</span></td>
                </tr>
            `;
        }).join("");

        $("#totalRegistros").textContent = state.dadosFiltrados.length;
    }

    function renderEmptyState(message) {
        elements.tableBody.innerHTML = `<tr class="empty-row"><td colspan="9">${escapeHtml(message)}</td></tr>`;
        $("#totalRegistros").textContent = 0;
    }

    function destroyChart(name) {
        if (state.charts[name]) {
            state.charts[name].destroy();
            state.charts[name] = null;
        }
    }

    function getDailyOccurrences() {
        const days = Array.from({ length: 31 }, (_, index) => String(index + 1).padStart(2, "0"));
        const total = state.dadosFiltrados.reduce((sum, item) => sum + getTotalAusencias(item), 0);
        const base = Math.floor(total / days.length);
        let rest = total % days.length;

        return days.map(() => {
            const value = base + (rest > 0 ? 1 : 0);
            rest -= 1;
            return value;
        });
    }

    function updateLineChart() {
        const canvas = $("#linha");
        const ctx = canvas.getContext("2d");
        destroyChart("linha");

        const gradient = ctx.createLinearGradient(0, 0, 0, 300);
        gradient.addColorStop(0, "rgba(79,111,82,.35)");
        gradient.addColorStop(1, "rgba(79,111,82,0)");

        const labels = Array.from({ length: 31 }, (_, index) => String(index + 1).padStart(2, "0"));
        state.charts.linha = new Chart(ctx, {
            type: "line",
            data: {
                labels,
                datasets: [{
                    label: "Ocorrências",
                    data: getDailyOccurrences(),
                    fill: true,
                    backgroundColor: gradient,
                    borderColor: COLORS.primary,
                    borderWidth: 3,
                    tension: .35,
                    pointRadius: 3,
                    pointHoverRadius: 7,
                    pointBackgroundColor: "#ffffff",
                    pointBorderColor: COLORS.primary,
                    pointBorderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: { y: { beginAtZero: true, ticks: { precision: 0 } } }
            }
        });
    }

    function updateDonutChart() {
        const container = $("#pizza");
        destroyChart("pizza");
        container.innerHTML = "";

        const totalFJ = state.dadosFiltrados.reduce((sum, item) => sum + getMetric(item, "FJ"), 0);
        const totalFN = state.dadosFiltrados.reduce((sum, item) => sum + getMetric(item, "FN"), 0);
        const totalA = state.dadosFiltrados.reduce((sum, item) => sum + getMetric(item, "A"), 0);
        const total = totalFJ + totalFN + totalA;

        state.charts.pizza = new ApexCharts(container, {
            chart: { type: "donut", height: 340, toolbar: { show: false } },
            series: [totalFJ, totalFN, totalA],
            labels: ["Falta Justificada", "Falta Não Justificada", "Atestado"],
            colors: [COLORS.warning, COLORS.danger, COLORS.info],
            legend: { position: "bottom", fontFamily: "Inter" },
            noData: { text: "Sem ocorrências" },
            plotOptions: {
                pie: {
                    donut: {
                        size: "72%",
                        labels: {
                            show: true,
                            total: {
                                show: true,
                                label: "Total",
                                formatter: () => total
                            }
                        }
                    }
                }
            }
        });

        state.charts.pizza.render();
    }

    function updateRankingChart() {
        const container = $("#ranking");
        destroyChart("ranking");
        container.innerHTML = "";

        const rankingData = [...state.dadosFiltrados]
            .map((item) => ({
                nome: item["NOME COMPLETO"],
                total: getTotalAusencias(item)
            }))
            .sort((a, b) => b.total - a.total)
            .slice(0, 10);

        state.charts.ranking = new ApexCharts(container, {
            chart: { type: "bar", height: 400, toolbar: { show: false } },
            plotOptions: {
                bar: {
                    horizontal: true,
                    borderRadius: 8,
                    distributed: true,
                    barHeight: "64%"
                }
            },
            colors: ["#3a5a40", "#4f6f52", "#588157", "#739072", "#8fb996", "#a7c957", "#b8d58a", "#cde8a1", "#d9ed92", "#e9f5bd"],
            series: [{
                name: "Ocorrências",
                data: rankingData.map((item) => item.total)
            }],
            xaxis: {
                categories: rankingData.map((item) => item.nome || "-")
            },
            legend: { show: false },
            grid: { borderColor: "#e5e7eb" },
            noData: { text: "Sem dados" }
        });

        state.charts.ranking.render();
    }

    function summarizeMonth(month) {
        const rows = state.workbookMonths[month]?.rows || [];
        return rows.reduce((sum, item) => sum + getTotalAusencias(item), 0);
    }

    function updateComparison() {
        const principal = state.activeMonth;
        const comparacao = elements.mesComparacao.value;
        const principalTotal = summarizeMonth(principal);
        const comparacaoTotal = comparacao ? summarizeMonth(comparacao) : null;

        const pills = [
            `<div class="compare-pill">${MONTH_LABELS[principal] || principal}<strong>${principalTotal}</strong></div>`
        ];

        if (comparacao) {
            const diff = principalTotal - comparacaoTotal;
            const signal = diff > 0 ? "+" : "";
            pills.push(`<div class="compare-pill">${MONTH_LABELS[comparacao] || comparacao}<strong>${comparacaoTotal}</strong></div>`);
            pills.push(`<div class="compare-pill">Diferença<strong>${signal}${diff}</strong></div>`);
        }

        elements.comparativoMeses.innerHTML = pills.join("");
    }

    function exportExcel() {
        if (!state.dadosFiltrados.length) {
            alert("Nenhum dado para exportar.");
            return;
        }

        const worksheet = XLSX.utils.json_to_sheet(state.dadosFiltrados);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Dashboard");
        XLSX.writeFile(workbook, `dashboard_sese_${state.activeMonth || "dados"}.xlsx`);
    }

    function saveOneDriveLink() {
        const link = elements.oneDriveLink.value.trim();
        if (!link) {
            alert("Cole um link válido do OneDrive.");
            return;
        }

        localStorage.setItem(LINK_KEY, link);
        setSyncStatus("Link salvo", "success");
        refreshData();
    }

    function startAutoRefresh(runNow = true) {
        stopAutoRefresh(false);
        state.isAutoRefreshActive = true;
        elements.btnAutoRefresh.classList.add("active");
        elements.btnAutoRefresh.title = "Desativar atualização automática";
        elements.btnAutoRefresh.querySelector("i").classList.add("rotating");
        elements.btnAutoRefresh.closest(".refresh-container")?.classList.add("active");
        state.autoRefreshInterval = setInterval(() => refreshData({ silent: true }), state.refreshIntervalTime);
        if (runNow) refreshData();
    }

    function stopAutoRefresh(updateUi = true) {
        if (state.autoRefreshInterval) {
            clearInterval(state.autoRefreshInterval);
            state.autoRefreshInterval = null;
        }

        state.isAutoRefreshActive = false;

        if (updateUi) {
            elements.btnAutoRefresh.classList.remove("active");
            elements.btnAutoRefresh.title = "Ativar atualização automática";
            elements.btnAutoRefresh.querySelector("i").classList.remove("rotating");
            elements.btnAutoRefresh.closest(".refresh-container")?.classList.remove("active");
        }
    }

    function bindEvents() {
        elements.toggleSidebar?.addEventListener("click", () => {
            if (window.innerWidth <= 992) {
                elements.sidebar.classList.toggle("active");
            } else {
                elements.sidebar.classList.toggle("collapsed");
                elements.main.classList.toggle("expanded");
            }
        });

        elements.themeBtn?.addEventListener("click", () => {
            document.body.classList.toggle("dark");
            elements.themeBtn.querySelector("i").className = document.body.classList.contains("dark")
                ? "fa-solid fa-sun"
                : "fa-solid fa-moon";
        });

        elements.btnUploadExcel?.addEventListener("click", () => elements.excelFile.click());
        elements.excelFile?.addEventListener("change", handleExcelUpload);
        elements.btnExportarExcel?.addEventListener("click", exportExcel);
        elements.btnExportPdf?.addEventListener("click", () => window.print());
        elements.btnFiltro?.addEventListener("click", applyFilters);
        elements.searchInput?.addEventListener("input", applyFilters);
        elements.turnoFilter?.addEventListener("change", applyFilters);
        elements.funcaoFilter?.addEventListener("change", applyFilters);

        elements.mesPrincipal?.addEventListener("change", () => setActiveMonth(elements.mesPrincipal.value));
        elements.mesComparacao?.addEventListener("change", updateComparison);
        elements.btnSalvarLink?.addEventListener("click", saveOneDriveLink);
        elements.oneDriveLink?.addEventListener("keydown", (event) => {
            if (event.key === "Enter") saveOneDriveLink();
        });

        elements.refreshInterval?.addEventListener("change", () => {
            state.refreshIntervalTime = Number(elements.refreshInterval.value);
            if (state.isAutoRefreshActive) startAutoRefresh();
        });

        elements.btnAutoRefresh?.addEventListener("click", () => {
            if (state.isAutoRefreshActive) {
                stopAutoRefresh();
                setSyncStatus("Pausado", "warning");
            } else {
                startAutoRefresh();
            }
        });
    }

    function init() {
        bindEvents();
        elements.oneDriveLink.value = localStorage.getItem(LINK_KEY) || DEFAULT_ONEDRIVE_LINK;
        state.refreshIntervalTime = Number(elements.refreshInterval.value || 30000);
        setSyncStatus("Iniciando", "neutral");
        refreshData().then(() => startAutoRefresh(false));
    }

    init();
});
