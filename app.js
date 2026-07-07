// ==========================================
// DASHBOARD SESE LOGÍSTICA - v5.0
// AUTO-REFRESH COM ONEDRIVE
// ==========================================

document.addEventListener("DOMContentLoaded", () => {

    // ==========================================
    // CONFIGURAÇÕES GLOBAIS
    // ==========================================

    const COLORS = {
        primary: "#4D7C0F",
        secondary: "#65A30D",
        success: "#16A34A",
        warning: "#F59E0B",
        danger: "#DC2626",
        dark: "#1F2937"
    };

    let dadosOriginais = [];
    let dadosFiltrados = [];

    let graficoLinha = null;
    let graficoPizza = null;
    let graficoRanking = null;

    let autoRefreshInterval = null;
    let isAutoRefreshActive = false;
    let refreshIntervalTime = 300000; // 5 minutos

    // Link do OneDrive
    const ONEDRIVE_LINK = "https://scaniaazureservices-my.sharepoint.com/:x:/r/personal/lucas_martins_scania_com/Documents/DASHBORD%20%20AB/Absente%C3%ADsmo%20-Expedi%C3%A7%C3%A3o%20Junho%202026.xlsx?d=w3a62a05b344044cd951d2eb4cd9d4280&csf=1&web=1&e=8kEhEM";

    // ==========================================
    // ELEMENTOS DO DOM
    // ==========================================

    const searchInput = document.getElementById("searchInput");
    const turnoFilter = document.getElementById("turnoFilter");
    const funcaoFilter = document.getElementById("funcaoFilter");
    const btnFiltro = document.getElementById("btnFiltro");

    const btnUploadExcel = document.getElementById("btnUploadExcel");
    const excelFile = document.getElementById("excelFile");
    const btnExportarExcel = document.getElementById("btnExportarExcel");

    const btnAutoRefresh = document.getElementById("btnAutoRefresh");
    const refreshInterval = document.getElementById("refreshInterval");
    const btnSalvarLink = document.getElementById("btnSalvarLink");
    const oneDriveLink = document.getElementById("oneDriveLink");

    const sidebar = document.querySelector(".sidebar");
    const toggleSidebar = document.querySelector(".toggle-sidebar");
    const main = document.querySelector(".main");
    const themeBtn = document.querySelector(".theme-btn");

    const mesPrincipal = document.getElementById("mesPrincipal");
    const mesComparacao = document.getElementById("mesComparacao");

    // ==========================================
    // SIDEBAR TOGGLE
    // ==========================================

    if (toggleSidebar) {
        toggleSidebar.addEventListener("click", () => {
            if (window.innerWidth <= 992) {
                sidebar.classList.toggle("active");
            } else {
                sidebar.classList.toggle("collapsed");
                main.classList.toggle("expanded");
            }
        });
    }

    // ==========================================
    // DARK MODE
    // ==========================================

    if (themeBtn) {
        themeBtn.addEventListener("click", () => {
            document.body.classList.toggle("dark");
            const icon = themeBtn.querySelector("i");
            icon.className = document.body.classList.contains("dark")
                ? "fa-solid fa-sun"
                : "fa-solid fa-moon";
        });
    }

    // ==========================================
    // IMPORTAR EXCEL
    // ==========================================

    if (btnUploadExcel) {
        btnUploadExcel.addEventListener("click", () => {
            excelFile.click();
        });
    }

    excelFile.addEventListener("change", carregarExcelLocal);

    function carregarExcelLocal(e) {
        const arquivo = e.target.files[0];
        if (!arquivo) return;

        const reader = new FileReader();
        reader.onload = function (evento) {
            const data = new Uint8Array(evento.target.result);
            const workbook = XLSX.read(data, { type: "array" });
            const primeiraAba = workbook.SheetNames[0];
            const planilha = workbook.Sheets[primeiraAba];

            dadosOriginais = XLSX.utils.sheet_to_json(planilha, { defval: "" });
            dadosFiltrados = [...dadosOriginais];

            popularFuncoes();
            atualizarDashboard();

            console.log("✅ Arquivo local carregado:", dadosOriginais.length, "registros");
        };
        reader.readAsArrayBuffer(arquivo);
    }

    // ==========================================
    // EXPORTAR EXCEL
    // ==========================================

    if (btnExportarExcel) {
        btnExportarExcel.addEventListener("click", () => {
            if (dadosFiltrados.length === 0) {
                alert("⚠️ Nenhum dado para exportar");
                return;
            }

            const ws = XLSX.utils.json_to_sheet(dadosFiltrados);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Dashboard");
            XLSX.writeFile(wb, "dashboard_sese.xlsx");
        });
    }

    // ==========================================
    // POPULAR FUNÇÕES
    // ==========================================

    function popularFuncoes() {
        const funcoes = [...new Set(dadosOriginais.map(x => x["Função"] || ""))];
        funcaoFilter.innerHTML = '<option value="">Todas</option>';
        funcoes.forEach(funcao => {
            if (funcao) funcaoFilter.innerHTML += `<option value="${funcao}">${funcao}</option>`;
        });
    }

    // ==========================================
    // FILTROS
    // ==========================================

    btnFiltro.addEventListener("click", aplicarFiltros);

    function aplicarFiltros() {
        const nome = searchInput.value.toLowerCase();
        const turno = turnoFilter.value;
        const funcao = funcaoFilter.value;

        dadosFiltrados = dadosOriginais.filter(item => {
            const nomeOk = !nome || (item["Nome"] && item["Nome"].toLowerCase().includes(nome));
            const turnoOk = !turno || item["Turno"] === turno;
            const funcaoOk = !funcao || item["Função"] === funcao;
            return nomeOk && turnoOk && funcaoOk;
        });

        atualizarDashboard();
    }

    // ==========================================
    // ATUALIZAR DASHBOARD
    // ==========================================

    function atualizarDashboard() {
        atualizarKPIs();
        atualizarTabela();
        atualizarGraficoLinha();
        atualizarGraficoPizza();
        atualizarRanking();
    }

    // ==========================================
    // KPI'S
    // ==========================================

    function atualizarKPIs() {
        const totalFuncionarios = dadosFiltrados.length;

        const totalFJ = dadosFiltrados.reduce((soma, item) => soma + Number(item["FJ"] || 0), 0);
        const totalFN = dadosFiltrados.reduce((soma, item) => soma + Number(item["FN"] || 0), 0);
        const totalAtestado = dadosFiltrados.reduce((soma, item) => soma + Number(item["Atestado"] || 0), 0);

        const totalFaltas = totalFJ + totalFN + totalAtestado;
        const presenca = totalFuncionarios > 0
            ? ((totalFuncionarios - totalFaltas) / totalFuncionarios) * 100
            : 100;

        document.getElementById("kpiFuncionarios").textContent = totalFuncionarios;
        document.getElementById("kpiFJ").textContent = totalFJ;
        document.getElementById("kpiFN").textContent = totalFN;
        document.getElementById("kpiPresenca").textContent = presenca.toFixed(1) + "%";
    }

    // ==========================================
    // TABELA
    // ==========================================

    function atualizarTabela() {
        const tbody = document.getElementById("tableBody");
        tbody.innerHTML = "";

        dadosFiltrados.forEach(item => {
            tbody.innerHTML += `
                <tr>
                    <td>${item["Nome"] || "-"}</td>
                    <td>${item["Função"] || "-"}</td>
                    <td>${item["Turno"] || "-"}</td>
                    <td><span class="badge badge-warning">${item["FJ"] || 0}</span></td>
                    <td><span class="badge badge-danger">${item["FN"] || 0}</span></td>
                    <td><span class="badge badge-success">${item["Atestado"] || 0}</span></td>
                </tr>
            `;
        });

        document.getElementById("totalRegistros").textContent = dadosFiltrados.length;
    }

    // ==========================================
    // GRÁFICO DE LINHA
    // ==========================================

    function atualizarGraficoLinha() {
        const ctx = document.getElementById("linha").getContext("2d");

        if (graficoLinha) graficoLinha.destroy();

        const gradient = ctx.createLinearGradient(0, 0, 0, 300);
        gradient.addColorStop(0, "rgba(77,124,15,.35)");
        gradient.addColorStop(1, "rgba(77,124,15,0)");

        const dias = Array.from({ length: 30 }, (_, i) => String(i + 1).padStart(2, "0"));
        const ocorrencias = dias.map(() => Math.floor(Math.random() * 4));

        graficoLinha = new Chart(ctx, {
            type: "line",
            data: {
                labels: dias,
                datasets: [{
                    label: "Ocorrências",
                    data: ocorrencias,
                    fill: true,
                    backgroundColor: gradient,
                    borderColor: COLORS.primary,
                    borderWidth: 3,
                    tension: 0.4,
                    pointRadius: 4,
                    pointHoverRadius: 8,
                    pointBackgroundColor: "#FFF",
                    pointBorderColor: COLORS.primary,
                    pointBorderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } }
            }
        });
    }

    // ==========================================
    // GRÁFICO PIZZA/DONUT
    // ==========================================

    function atualizarGraficoPizza() {
        const container = document.querySelector("#pizza");
        container.innerHTML = "";

        const totalFJ = dadosFiltrados.reduce((soma, item) => soma + Number(item["FJ"] || 0), 0);
        const totalFN = dadosFiltrados.reduce((soma, item) => soma + Number(item["FN"] || 0), 0);
        const totalAtestado = dadosFiltrados.reduce((soma, item) => soma + Number(item["Atestado"] || 0), 0);
        const total = totalFJ + totalFN + totalAtestado;

        graficoPizza = new ApexCharts(container, {
            chart: { type: "donut", height: 350, toolbar: { show: false } },
            series: [totalFJ, totalFN, totalAtestado],
            labels: ["FJ", "FN", "Atestado"],
            colors: [COLORS.warning, COLORS.danger, COLORS.success],
            legend: { position: "bottom", fontFamily: "Inter" },
            plotOptions: {
                pie: {
                    donut: {
                        size: "72%",
                        labels: {
                            show: true,
                            total: { show: true, label: "Total", formatter: () => total }
                        }
                    }
                }
            }
        });

        graficoPizza.render();
    }

    // ==========================================
    // TOP 10 COLABORADORES
    // ==========================================

    function atualizarRanking() {
        const rankingDiv = document.querySelector("#ranking");
        rankingDiv.innerHTML = "";

        const rankingDados = [...dadosFiltrados]
            .sort((a, b) => {
                const totalA = Number(a.FJ || 0) + Number(a.FN || 0);
                const totalB = Number(b.FJ || 0) + Number(b.FN || 0);
                return totalB - totalA;
            })
            .slice(0, 10);

        graficoRanking = new ApexCharts(rankingDiv, {
            chart: { type: "bar", height: 400, toolbar: { show: false } },
            plotOptions: {
                bar: { horizontal: true, borderRadius: 10, distributed: true, barHeight: "65%" }
            },
            colors: ["#4D7C0F", "#5B8C14", "#6B9E1C", "#7CB342", "#8BC34A", "#9CCC65", "#AED581", "#C5E1A5", "#DCE775", "#E6EE9C"],
            series: [{
                name: "Faltas",
                data: rankingDados.map(item => Number(item.FJ || 0) + Number(item.FN || 0))
            }],
            xaxis: {
                categories: rankingDados.map(item => item.Nome)
            },
            legend: { show: false },
            grid: { borderColor: "#E5E7EB" }
        });

        graficoRanking.render();
    }

    // ==========================================
    // COMPARAÇÃO ENTRE MESES
    // ==========================================

    if (mesPrincipal && mesComparacao) {
        mesPrincipal.addEventListener("change", () => {
            console.log("Mês principal alterado para:", mesPrincipal.value);
        });
        mesComparacao.addEventListener("change", () => {
            console.log("Mês de comparação alterado para:", mesComparacao.value);
        });
    }

    // ==========================================
    // SINCRONIZAÇÃO ONEDRIVE
    // ==========================================

    function converterLinkParaDownload(link) {
        if (!link) return null;

        let urlLimpa = link.trim();

        if (urlLimpa.includes('sharepoint.com')) {
            if (urlLimpa.includes('?')) {
                const baseUrl = urlLimpa.split('?')[0];
                const dParam = urlLimpa.match(/[?&]d=([^&]+)/);

                if (dParam && dParam[1]) {
                    return `${baseUrl}?download=1&d=${dParam[1]}`;
                }
            }

            if (!urlLimpa.includes('download=1')) {
                urlLimpa += (urlLimpa.includes('?') ? '&' : '?') + 'download=1';
            }
        }

        return urlLimpa;
    }

    function salvarLinkOneDrive() {
        const link = oneDriveLink.value.trim();

        if (!link) {
            alert("⚠️ Cole um link válido do OneDrive");
            return;
        }

        localStorage.setItem('oneDriveLink', link);
        console.log("✅ Link do OneDrive salvo");
        alert("✅ Link salvo com sucesso!");
    }

    function carregarLinkOneDrive() {
        let link = localStorage.getItem('oneDriveLink');

        if (!link) {
            link = ONEDRIVE_LINK;
            localStorage.setItem('oneDriveLink', link);
        }

        if (link && oneDriveLink) {
            oneDriveLink.value = link;
            console.log("✅ Link carregado");
        }
    }

    async function baixarDoOneDrive() {
        try {
            const link = localStorage.getItem('oneDriveLink');

            if (!link) {
                const syncStatus = document.getElementById("syncStatus");
                if (syncStatus) {
                    syncStatus.textContent = "❌ Sem link";
                    syncStatus.style.color = "#dc2626";
                }
                return;
            }

            const container = document.querySelector(".refresh-container");
            const syncStatus = document.getElementById("syncStatus");

            if (syncStatus) {
                syncStatus.textContent = "🔄 Sincronizando...";
                container.classList.add("syncing");
            }

            console.log("📥 Baixando arquivo do OneDrive...");

            const downloadUrl = converterLinkParaDownload(link);
            console.log("URL:", downloadUrl);

            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 60000);

            const response = await fetch(downloadUrl, {
                method: 'GET',
                signal: controller.signal,
                headers: { 'Cache-Control': 'no-cache' }
            });

            clearTimeout(timeoutId);

            if (!response.ok) {
                throw new Error(`HTTP ${response.status}`);
            }

            const arrayBuffer = await response.arrayBuffer();
            console.log("Tamanho:", arrayBuffer.byteLength, "bytes");

            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, { type: "array" });

            console.log("Abas:", workbook.SheetNames);

            const planilha = workbook.Sheets[workbook.SheetNames[0]];
            dadosOriginais = XLSX.utils.sheet_to_json(planilha, { defval: "" });
            dadosFiltrados = [...dadosOriginais];

            console.log("Dados:", dadosOriginais.length, "registros");

            popularFuncoes();
            atualizarDashboard();

            const agora = new Date().toLocaleTimeString("pt-BR");

            if (syncStatus) {
                syncStatus.textContent = `✅ ${agora}`;
                syncStatus.style.color = "#16A34A";
                container.classList.remove("syncing");
            }

            console.log(`✅ Sincronizado: ${dadosOriginais.length} colaboradores`);

        } catch (erro) {
            const syncStatus = document.getElementById("syncStatus");
            if (syncStatus) {
                syncStatus.textContent = "❌ Erro";
                syncStatus.style.color = "#dc2626";
            }
            console.error("❌ Erro:", erro.message);
        }
    }

    function iniciarAutoRefresh() {
        isAutoRefreshActive = true;
        btnAutoRefresh.classList.add("active");

        const container = btnAutoRefresh.closest(".refresh-container");
        if (container) container.classList.add("active");

        btnAutoRefresh.title = "Desativar Auto-Atualização";

        const icon = btnAutoRefresh.querySelector("i");
        icon.classList.add("rotating");

        baixarDoOneDrive();

        autoRefreshInterval = setInterval(() => {
            console.log("🔄 Atualizando...");
            baixarDoOneDrive();
        }, refreshIntervalTime);

        console.log(`✅ Auto-refresh: ${refreshIntervalTime / 1000}s`);
    }

    function pararAutoRefresh() {
        isAutoRefreshActive = false;
        btnAutoRefresh.classList.remove("active");

        const container = btnAutoRefresh.closest(".refresh-container");
        if (container) container.classList.remove("active");

        btnAutoRefresh.title = "Ativar Auto-Atualização";

        const icon = btnAutoRefresh.querySelector("i");
        icon.classList.remove("rotating");

        if (autoRefreshInterval) {
            clearInterval(autoRefreshInterval);
            autoRefreshInterval = null;
        }

        console.log("⏹️ Auto-refresh desativado");
    }

    function mudarIntervalo() {
        refreshIntervalTime = parseInt(refreshInterval.value);

        if (isAutoRefreshActive) {
            pararAutoRefresh();
            iniciarAutoRefresh();
        }
    }

    // ==========================================
    // EVENT LISTENERS
    // ==========================================

    if (btnAutoRefresh) {
        btnAutoRefresh.addEventListener("click", () => {
            if (isAutoRefreshActive) {
                pararAutoRefresh();
            } else {
                iniciarAutoRefresh();
            }
        });
    }

    if (refreshInterval) {
        refreshInterval.addEventListener("change", mudarIntervalo);
    }

    if (btnSalvarLink) {
        btnSalvarLink.addEventListener("click", salvarLinkOneDrive);
    }

    if (oneDriveLink) {
        oneDriveLink.addEventListener("keypress", (e) => {
            if (e.key === "Enter") salvarLinkOneDrive();
        });
    }

    // ==========================================
    // INICIALIZAÇÃO
    // ==========================================

    carregarLinkOneDrive();
    popularFuncoes();
    atualizarDashboard();

    // Inicia auto-refresh após 1 segundo
    setTimeout(() => {
        console.log("🚀 Iniciando sincronização automática...");
        iniciarAutoRefresh();
    }, 1000);

});

// ==========================================
// FIM APP.JS
// ==========================================
