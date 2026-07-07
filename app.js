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
    const ONEDRIVE_LINK = "https://scaniaazureservices-my.sharepoint.com/:x:/r/personal/lucas_martins_scania_com/_layouts/15/Doc.aspx?sourcedoc=%7B3A62A05B-3440-44CD-951D-2EB4CD9D4280%7D&file=Absente%25u00edsmo%20-Expedi%25u00e7%25u00e3o%20Junho%202026.xlsx&openShare=true&fromShare=true&action=default&mobileredirect=true;

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
            
            // Buscar a aba JUNHO ou a primeira aba com dados
            let nomeAba = "JUNHO";
            if (!workbook.SheetNames.includes("JUNHO")) {
                nomeAba = workbook.SheetNames.find(aba => aba !== "RESUMO");
                if (!nomeAba) nomeAba = workbook.SheetNames[0];
            }
            
            const planilha = workbook.Sheets[nomeAba];
            
            // Ler os dados, pulando título e cabeçalho
            const dados = XLSX.utils.sheet_to_json(planilha, { header: 1 });
            
            // Encontrar linha do cabeçalho (que é "TURNO", "NOME COMPLETO", etc)
            let linhaHeader = 0;
            for (let i = 0; i < dados.length; i++) {
                if (dados[i][0] === "TURNO") {
                    linhaHeader = i;
                    break;
                }
            }
            
            // Extrair headers e dados
            const headers = dados[linhaHeader];
            dadosOriginais = [];
            
            for (let i = linhaHeader + 1; i < dados.length; i++) {
                if (!dados[i][0]) continue; // Pular linhas vazias
                
                const obj = {};
                headers.forEach((header, idx) => {
                    obj[header] = dados[i][idx] || "";
                });
                dadosOriginais.push(obj);
            }
            
            dadosFiltrados = [...dadosOriginais];

            popularFuncoes();
            atualizarDashboard();

            console.log("✅ Arquivo carregado de:", nomeAba, "com", dadosOriginais.length, "registros");
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
        const funcoes = [...new Set(dadosOriginais.map(x => x["FUNÇÃO"] || "").filter(Boolean))];
        funcaoFilter.innerHTML = '<option value="">Todas</option>';
        funcoes.forEach(funcao => {
            funcaoFilter.innerHTML += `<option value="${funcao}">${funcao}</option>`;
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
            const nomeOk = !nome || (item["NOME COMPLETO"] && item["NOME COMPLETO"].toLowerCase().includes(nome));
            const turnoOk = !turno || item["TURNO"] === turno;
            const funcaoOk = !funcao || item["FUNÇÃO"] === funcao;
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
        const totalA = dadosFiltrados.reduce((soma, item) => soma + Number(item["A"] || 0), 0);

        const totalFaltas = totalFJ + totalFN + totalA;
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
                    <td>${item["NOME COMPLETO"] || "-"}</td>
                    <td>${item["FUNÇÃO"] || "-"}</td>
                    <td>${item["TURNO"] || "-"}</td>
                    <td><span class="badge badge-warning">${item["FJ"] || 0}</span></td>
                    <td><span class="badge badge-danger">${item["FN"] || 0}</span></td>
                    <td><span class="badge badge-info">${item["A"] || 0}</span></td>
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
        const totalA = dadosFiltrados.reduce((soma, item) => soma + Number(item["A"] || 0), 0);
        const total = totalFJ + totalFN + totalA;

        graficoPizza = new ApexCharts(container, {
            chart: { type: "donut", height: 350, toolbar: { show: false } },
            series: [totalFJ, totalFN, totalA],
            labels: ["Feriado Judaico", "Feriado Nacional", "Ausência"],
            colors: [COLORS.warning, COLORS.danger, COLORS.primary],
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
                const totalA = Number(a.FJ || 0) + Number(a.FN || 0) + Number(a.A || 0);
                const totalB = Number(b.FJ || 0) + Number(b.FN || 0) + Number(b.A || 0);
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
                data: rankingDados.map(item => Number(item.FJ || 0) + Number(item.FN || 0) + Number(item.A || 0))
            }],
            xaxis: {
                categories: rankingDados.map(item => item["NOME COMPLETO"])
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
                console.warn("⚠️ Nenhum link configurado");
                return Promise.reject("Sem link");
            }

            const container = document.querySelector(".refresh-container");
            const syncStatus = document.getElementById("syncStatus");

            if (syncStatus) {
                syncStatus.textContent = "🔄 Sincronizando...";
                if (container) container.classList.add("syncing");
            }

            console.log("📥 Tentando baixar de:", link.substring(0, 50) + "...");

            const downloadUrl = converterLinkParaDownload(link);

            // Tenta com fetch direto
            let response = await Promise.race([
                fetch(downloadUrl, {
                    method: 'GET',
                    headers: { 'Cache-Control': 'no-cache' }
                }),
                new Promise((_, reject) => 
                    setTimeout(() => reject(new Error('Timeout')), 15000)
                )
            ]).catch(err => {
                console.warn("⚠️ Fetch direto falhou:", err.message);
                return null;
            });

            if (!response || !response.ok) {
                console.error("❌ Falha ao baixar do OneDrive");
                
                if (syncStatus) {
                    syncStatus.textContent = "❌ Erro";
                    syncStatus.style.color = "#dc2626";
                    if (container) container.classList.remove("syncing");
                }

                throw new Error("SharePoint retornou erro. Tente importar manualmente.");
            }

            const arrayBuffer = await response.arrayBuffer();
            const data = new Uint8Array(arrayBuffer);
            const workbook = XLSX.read(data, { type: "array" });

            // Buscar a aba JUNHO
            let nomeAba = "JUNHO";
            if (!workbook.SheetNames.includes("JUNHO")) {
                nomeAba = workbook.SheetNames.find(aba => aba !== "RESUMO");
                if (!nomeAba) nomeAba = workbook.SheetNames[0];
            }

            const planilha = workbook.Sheets[nomeAba];
            
            // Ler os dados corretamente
            const dados = XLSX.utils.sheet_to_json(planilha, { header: 1 });
            
            let linhaHeader = 0;
            for (let i = 0; i < dados.length; i++) {
                if (dados[i][0] === "TURNO") {
                    linhaHeader = i;
                    break;
                }
            }
            
            const headers = dados[linhaHeader];
            dadosOriginais = [];
            
            for (let i = linhaHeader + 1; i < dados.length; i++) {
                if (!dados[i][0]) continue;
                
                const obj = {};
                headers.forEach((header, idx) => {
                    obj[header] = dados[i][idx] || "";
                });
                dadosOriginais.push(obj);
            }

            dadosFiltrados = [...dadosOriginais];

            console.log("✅ Dados carregados:", dadosOriginais.length, "registros de", nomeAba);

            popularFuncoes();
            atualizarDashboard();

            const agora = new Date().toLocaleTimeString("pt-BR");

            if (syncStatus) {
                syncStatus.textContent = `✅ ${agora}`;
                syncStatus.style.color = "#16A34A";
                if (container) container.classList.remove("syncing");
            }
            
            return Promise.resolve();

        } catch (erro) {
            const syncStatus = document.getElementById("syncStatus");
            const container = document.querySelector(".refresh-container");
            
            if (syncStatus) {
                syncStatus.textContent = "❌ Offline";
                syncStatus.style.color = "#dc2626";
                if (container) container.classList.remove("syncing");
            }

            console.error("❌ Erro de sincronização:", erro.message);
            return Promise.reject(erro);
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

        // Se tiver dados, inicia a sincronização
        if (dadosOriginais.length > 0) {
            baixarDoOneDrive();
        } else {
            console.log("💡 Importe um arquivo Excel primeiro");
        }

        autoRefreshInterval = setInterval(() => {
            console.log("🔄 Atualizando...");
            
            // Tenta sincronizar com OneDrive
            baixarDoOneDrive().catch(err => {
                // Se falhar, apenas recarrega os dados atuais
                if (dadosOriginais.length > 0) {
                    atualizarDashboard();
                    console.log("🔄 Dashboard recarregado (dados locais)");
                }
            });
        }, refreshIntervalTime);

        console.log(`✅ Auto-refresh: ${refreshIntervalTime / 1000 / 60}min`);
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
    
    // Tentar carregar dados do OneDrive automaticamente
    console.log("📥 Tentando carregar arquivo do OneDrive...");
    baixarDoOneDrive().then(() => {
        // Se conseguir baixar, inicia o auto-refresh a cada 5 minutos
        console.log("✅ Auto-refresh do OneDrive iniciado");
        iniciarAutoRefresh();
    }).catch(() => {
        // Se não conseguir, espera o usuário importar manualmente
        console.log("⚠️ Não foi possível carregar do OneDrive. Importe um arquivo Excel manualmente.");
        popularFuncoes();
        atualizarDashboard();
    });

    console.log("✅ Dashboard carregado");
    console.log("📋 Aguardando sincronização ou importação manual...");

});

// ==========================================
// FIM APP.JS
// ==========================================
