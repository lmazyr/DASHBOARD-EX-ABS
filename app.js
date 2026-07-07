// ==========================================
// DASHBOARD SESE LOGÍSTICA - v4.0
// APP.JS
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
    let refreshIntervalTime = 30000;

    const ONEDRIVE_URL = "https://scaniaazureservices-my.sharepoint.com/:x:/g/personal/lucas_martins_scania_com/IQBboGI6QDTNRJUdLrTNnUKAAXhfsDUEmnLnBmsrURk6FIc";

    // ==========================================
    // ELEMENTOS
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

    const themeBtn = document.querySelector(".theme-btn");

    // ==========================================
    // SIDEBAR
    // ==========================================

    const sidebar = document.querySelector(".sidebar");
    const toggleSidebar = document.querySelector(".toggle-sidebar");
    const main = document.querySelector(".main");

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

    excelFile.addEventListener("change", carregarExcel);

    function carregarExcel(e) {

        const arquivo = e.target.files[0];

        if (!arquivo) return;

        const reader = new FileReader();

        reader.onload = function (evento) {

            const data = new Uint8Array(evento.target.result);

            const workbook = XLSX.read(data, {
                type: "array"
            });

            const primeiraAba = workbook.SheetNames[0];

            const planilha = workbook.Sheets[primeiraAba];

            dadosOriginais = XLSX.utils.sheet_to_json(planilha);

            dadosFiltrados = [...dadosOriginais];

            popularFuncoes();

            atualizarDashboard();

        };

        reader.readAsArrayBuffer(arquivo);

    }

    // ==========================================
    // POPULAR FUNÇÕES
    // ==========================================

    function popularFuncoes() {

        const funcoes = [
            ...new Set(
                dadosOriginais.map(x => x["Função"])
            )
        ];

        funcaoFilter.innerHTML = `
            <option value="">
                Todas
            </option>
        `;

        funcoes.forEach(funcao => {

            funcaoFilter.innerHTML += `
                <option value="${funcao}">
                    ${funcao}
                </option>
            `;

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

            const nomeOk =
                !nome ||
                item["Nome"]
                    ?.toLowerCase()
                    .includes(nome);

            const turnoOk =
                !turno ||
                item["Turno"] === turno;

            const funcaoOk =
                !funcao ||
                item["Função"] === funcao;

            return (
                nomeOk &&
                turnoOk &&
                funcaoOk
            );

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

        const totalFJ = dadosFiltrados.reduce((soma, item) => {

            return soma + Number(item["FJ"] || 0);

        }, 0);

        const totalFN = dadosFiltrados.reduce((soma, item) => {

            return soma + Number(item["FN"] || 0);

        }, 0);

        const totalAtestado = dadosFiltrados.reduce((soma, item) => {

            return soma + Number(item["Atestado"] || 0);

        }, 0);

        const totalFaltas =
            totalFJ +
            totalFN +
            totalAtestado;

        const presenca = totalFuncionarios > 0
            ? (
                (
                    totalFuncionarios - totalFaltas
                ) /
                totalFuncionarios
            ) * 100
            : 100;

        document.getElementById(
            "kpiFuncionarios"
        ).textContent = totalFuncionarios;

        document.getElementById(
            "kpiFJ"
        ).textContent = totalFJ;

        document.getElementById(
            "kpiFN"
        ).textContent = totalFN;

        document.getElementById(
            "kpiPresenca"
        ).textContent =
            presenca.toFixed(1) + "%";

    }

    // ==========================================
    // TABELA
    // ==========================================

    function atualizarTabela() {

        const tbody =
            document.getElementById(
                "tableBody"
            );

        tbody.innerHTML = "";

        dadosFiltrados.forEach(item => {

            tbody.innerHTML += `

                <tr>

                    <td>${item["Nome"] || "-"}</td>

                    <td>${item["Função"] || "-"}</td>

                    <td>${item["Turno"] || "-"}</td>

                    <td>

                        <span class="badge badge-warning">

                            ${item["FJ"] || 0}

                        </span>

                    </td>

                    <td>

                        <span class="badge badge-danger">

                            ${item["FN"] || 0}

                        </span>

                    </td>

                    <td>

                        <span class="badge badge-success">

                            ${item["Atestado"] || 0}

                        </span>

                    </td>

                </tr>

            `;

        });

        document.getElementById(
            "totalRegistros"
        ).textContent =
            dadosFiltrados.length;

    }

    // ==========================================
    // EXPORTAR EXCEL
    // ==========================================

    function exportarExcel() {

        if (dadosFiltrados.length === 0) {

            alert("⚠️ Nenhum dado para exportar");

            return;

        }

        const ws = XLSX.utils.json_to_sheet(dadosFiltrados);

        const wb = XLSX.utils.book_new();

        XLSX.utils.book_append_sheet(wb, ws, "Dashboard");

        XLSX.writeFile(wb, "dashboard_sese.xlsx");

    }

    if (btnExportarExcel) {

        btnExportarExcel.addEventListener("click", exportarExcel);

    }
   


    // ==========================================
    // GRÁFICO DE LINHA
    // ==========================================

    function atualizarGraficoLinha() {

        const ctx =
            document
                .getElementById("linha")
                .getContext("2d");

        if (graficoLinha) {

            graficoLinha.destroy();

        }

        const gradient =
            ctx.createLinearGradient(
                0,
                0,
                0,
                300
            );

        gradient.addColorStop(
            0,
            "rgba(77,124,15,.35)"
        );

        gradient.addColorStop(
            1,
            "rgba(77,124,15,0)"
        );

        const dias = [
            "01","02","03","04","05",
            "06","07","08","09","10",
            "11","12","13","14","15",
            "16","17","18","19","20",
            "21","22","23","24","25",
            "26","27","28","29","30"
        ];

        const ocorrencias = dias.map(() => {

            return Math.floor(
                Math.random() * 4
            );

        });

        graficoLinha = new Chart(ctx, {

            type: "line",

            data: {

                labels: dias,

                datasets: [{

                    label: "Ocorrências",

                    data: ocorrencias,

                    fill: true,

                    backgroundColor:
                        gradient,

                    borderColor:
                        COLORS.primary,

                    borderWidth: 3,

                    tension: .4,

                    pointRadius: 4,

                    pointHoverRadius: 8,

                    pointBackgroundColor:
                        "#FFF",

                    pointBorderColor:
                        COLORS.primary,

                    pointBorderWidth: 2

                }]

            },

            options: {

                responsive: true,

                maintainAspectRatio: false,

                plugins: {

                    legend: {

                        display: false

                    }

                },

                scales: {

                    y: {

                        beginAtZero: true,

                        ticks: {

                            stepSize: 1

                        }

                    }

                }

            }

        });

    }
    
    // ==========================================
    // GRÁFICO DONUT
    // ==========================================

    function atualizarGraficoPizza() {

        const container = 
            document.querySelector("#pizza");

        container.innerHTML = "";

        const totalFJ = dadosFiltrados.reduce(
            (soma, item) =>
                soma + Number(item["FJ"] || 0),
            0
        );

        const totalFN = dadosFiltrados.reduce(
            (soma, item) =>
                soma + Number(item["FN"] || 0),
            0
        );

        const totalAtestado = dadosFiltrados.reduce(
            (soma, item) =>
                soma + Number(item["Atestado"] || 0),
            0
        );

        const total =
            totalFJ +
            totalFN +
            totalAtestado;

        graficoPizza = new ApexCharts(
            container,
            {

                chart: {

                    type: "donut",
                    height: 350,

                    toolbar: {
                        show: false
                    }

                },

                series: [

                    totalFJ,
                    totalFN,
                    totalAtestado

                ],

                labels: [

                    "FJ",
                    "FN",
                    "Atestado"

                ],

                colors: [

                    COLORS.warning,
                    COLORS.danger,
                    COLORS.success

                ],

                legend: {

                    position: "bottom",

                    fontFamily: "Inter"

                },

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

            }

        );

        graficoPizza.render();

    }

    // ==========================================
    // TOP 10 COLABORADORES
    // ==========================================

    function atualizarRanking() {

        const rankingDiv =
            document.querySelector("#ranking");

        rankingDiv.innerHTML = "";

        const rankingDados =
            [...dadosFiltrados]

                .sort((a, b) => {

                    const totalA =
                        Number(a.FJ || 0) +
                        Number(a.FN || 0);

                    const totalB =
                        Number(b.FJ || 0) +
                        Number(b.FN || 0);

                    return totalB - totalA;

                })

                .slice(0, 10);

        graficoRanking = new ApexCharts(

            rankingDiv,

            {

                chart: {

                    type: "bar",

                    height: 400,

                    toolbar: {

                        show: false

                    }

                },

                plotOptions: {

                    bar: {

                        horizontal: true,

                        borderRadius: 10,

                        distributed: true,

                        barHeight: "65%"

                    }

                },

                colors: [

                    "#4D7C0F",
                    "#5B8C14",
                    "#6B9E1C",
                    "#7CB342",
                    "#8BC34A",
                    "#9CCC65",
                    "#AED581",
                    "#C5E1A5",
                    "#DCE775",
                    "#E6EE9C"

                ],

                series: [{

                    name: "Faltas",

                    data:

                        rankingDados.map(

                            item =>

                                Number(item.FJ || 0) +

                                Number(item.FN || 0)

                        )

                }],

                xaxis: {

                    categories:

                        rankingDados.map(

                            item => item.Nome

                        )

                },

                legend: {

                    show: false

                },

                grid: {

                    borderColor: "#E5E7EB"

                }

            }

        );

        graficoRanking.render();

    }

    // ==========================================
    // COMPARAÇÃO ENTRE MESES
    // ==========================================

    const mesPrincipal =
        document.getElementById("mesPrincipal");

    const mesComparacao =
        document.getElementById("mesComparacao");

    if (mesPrincipal && mesComparacao) {

        mesPrincipal.addEventListener(

            "change",

            atualizarComparacao

        );

        mesComparacao.addEventListener(

            "change",

            atualizarComparacao

        );

    }

    function atualizarComparacao() {

        const mes1 =
            mesPrincipal.value;

        const mes2 =
            mesComparacao.value;

        console.log(

            `Comparando ${mes1} x ${mes2}`

        );

        // Futuramente:
        // Atualizar gráficos usando os dados
        // dos dois meses importados

    }

    // ==========================================
    // AUTO-REFRESH
    // ==========================================

    async function baixarDoOneDrive() {

        try {

            const container = document.querySelector(".refresh-container");

            const syncStatus = document.getElementById("syncStatus");

            if (syncStatus) {

                syncStatus.textContent = "🔄 Sincronizando...";

                container.classList.add("syncing");

            }

            console.log("📥 Baixando arquivo do OneDrive...");

            const downloadUrl = ONEDRIVE_URL.replace('/view?', '/download?');

            const response = await fetch(downloadUrl);

            if (!response.ok) {

                throw new Error(`Erro ao baixar: ${response.status}`);

            }

            const arrayBuffer = await response.arrayBuffer();

            const data = new Uint8Array(arrayBuffer);

            const workbook = XLSX.read(data, { type: "array" });

            const primeiraAba = workbook.SheetNames[0];

            const planilha = workbook.Sheets[primeiraAba];

            dadosOriginais = XLSX.utils.sheet_to_json(planilha);

            dadosFiltrados = [...dadosOriginais];

            popularFuncoes();

            atualizarDashboard();

            const agora = new Date().toLocaleTimeString("pt-BR");

            if (syncStatus) {

                syncStatus.textContent = `✅ ${agora}`;

                container.classList.remove("syncing");

            }

            console.log(`✅ Arquivo sincronizado: ${dadosOriginais.length} colaboradores`);

        } catch (erro) {

            const syncStatus = document.getElementById("syncStatus");

            if (syncStatus) {

                syncStatus.textContent = "❌ Erro ao sincronizar";

                syncStatus.style.color = "#dc2626";

            }

            console.error("❌ Erro ao sincronizar com OneDrive:", erro);

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

            console.log("🔄 Sincronizando com OneDrive...");

            baixarDoOneDrive();

        }, refreshIntervalTime);

        console.log(`✅ Auto-refresh ativado a cada ${refreshIntervalTime / 1000} segundos`);

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

            console.log(`✅ Intervalo alterado para ${refreshIntervalTime / 1000} segundos`);

        }

    }

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

    

    popularFuncoes();

    atualizarDashboard();

});

// ==========================================
// FIM APP.JS
// ==========================================
