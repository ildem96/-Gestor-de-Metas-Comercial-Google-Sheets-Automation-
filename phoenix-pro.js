function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("📊 Relatórios");

    menu.addItem("🔄 Atualizar Dashboard Geral", "atualizarDashboardCompleto");
    menu.addSeparator();

    // Busca dinâmica dos vendedores baseada na coluna J (Coluna 10)
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const abaDados = ss.getSheetByName("Banco de Dados");

        if (abaDados) {
            const ultimaLinha = abaDados.getLastRow();
            if (ultimaLinha > 1) {
                const limiteBusca = Math.min(ultimaLinha - 1, 500);
                const linhaInicio = ultimaLinha - limiteBusca + 1;

                const dadosVendedores = abaDados.getRange(linhaInicio, 10, limiteBusca, 1).getValues();
                const vendedoresUnicos = [
                    ...new Set(dadosVendedores.map((v) => v[0].toString().trim()).filter((v) => v !== ""))
                ].sort();

                PropertiesService.getDocumentProperties().setProperty(
                    "listaVendedores",
                    JSON.stringify(vendedoresUnicos)
                );

                vendedoresUnicos.forEach((vendedor, index) => {
                    if (index < 10) {
                        menu.addItem("👤 " + vendedor, "abrirVendedor" + index);
                    }
                });
            }
        }
    } catch (e) {
        // Silencia erros caso a planilha seja aberta sem todas as permissões carregadas
    }

    menu.addToUi();
}

// ====================================================================
// GATILHOS EXCLUSIVOS PARA OS VENDEDORES DO MENU
// ====================================================================
function abrirVendedor0() {
    abrirDashboardPorVendedorIndex(0);
}
function abrirVendedor1() {
    abrirDashboardPorVendedorIndex(1);
}
function abrirVendedor2() {
    abrirDashboardPorVendedorIndex(2);
}
function abrirVendedor3() {
    abrirDashboardPorVendedorIndex(3);
}
function abrirVendedor4() {
    abrirDashboardPorVendedorIndex(4);
}
function abrirVendedor5() {
    abrirDashboardPorVendedorIndex(5);
}
function abrirVendedor6() {
    abrirDashboardPorVendedorIndex(6);
}
function abrirVendedor7() {
    abrirDashboardPorVendedorIndex(7);
}
function abrirVendedor8() {
    abrirDashboardPorVendedorIndex(8);
}
function abrirVendedor9() {
    abrirDashboardPorVendedorIndex(9);
}

function abrirDashboardPorVendedorIndex(index) {
    const propriedades = PropertiesService.getDocumentProperties().getProperty("listaVendedores");
    if (propriedades) {
        const vendedores = JSON.parse(propriedades);
        if (vendedores[index]) {
            abrirDashboardHTML(vendedores[index]);
        }
    } else {
        SpreadsheetApp.getUi().alert(
            "⚠️ Lista de vendedores não encontrada. Atualize a página (F5) e tente novamente."
        );
    }
}

// ====================================================================
// FUNÇÕES UTILITÁRIAS (AQUI ESTÁ A FUNÇÃO CORRIGIDA)
// ====================================================================
function getInicioDaSemana(data) {
    let d = new Date(data);
    let day = d.getDay();
    // Ajusta para segunda-feira como o primeiro dia da semana comercial
    let diff = d.getDate() - day + (day == 0 ? -6 : 1);
    return new Date(d.setDate(diff)).setHours(0, 0, 0, 0);
}

function obterComponentesData(valor, timeZone) {
    if (valor instanceof Date && !isNaN(valor.getTime())) {
        let dia = Utilities.formatDate(valor, timeZone, "dd");
        let mes = Utilities.formatDate(valor, timeZone, "MM");
        let ano = Utilities.formatDate(valor, timeZone, "yyyy");
        return { dia: parseInt(dia, 10), mes: parseInt(mes, 10) - 1, ano: parseInt(ano, 10), valida: true };
    }
    if (valor !== null && valor !== undefined) {
        let str = valor.toString().trim();
        let partes = str.split("/");
        if (partes.length === 3) {
            let dia = parseInt(partes[0], 10);
            let mes = parseInt(partes[1], 10) - 1;
            let ano = parseInt(partes[2].length === 2 ? "20" + partes[2] : partes[2], 10);
            return { dia: dia, mes: mes, ano: ano, valida: !isNaN(dia) && !isNaN(mes) && !isNaN(ano) };
        }
    }
    return { valida: false };
}

// ====================================================================
// DASHBOARD WEB DO VENDEDOR (HTML + CHART.JS) - MARGENS CORRIGIDAS
// ====================================================================
function abrirDashboardHTML(vendedorSelecionado) {
    if (!vendedorSelecionado) vendedorSelecionado = "Visualização de Teste";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaDados = ss.getSheetByName("Banco de Dados");
    const abaDash = ss.getSheetByName("DashBoard");
    const tZ = ss.getSpreadsheetTimeZone();

    const ultimaLinha = abaDados.getLastRow();
    if (ultimaLinha < 2) return;

    const maxLinhas = 5000;
    const linhasParaLer = Math.min(ultimaLinha - 1, maxLinhas);
    const linhaInicial = ultimaLinha - linhasParaLer + 1;

    let dados = abaDados.getRange(linhaInicial, 1, linhasParaLer, 10).getValues();

    const metaReal = abaDash.getRange("B2").getValue() || 1;
    const metaHipotetica = abaDash.getRange("C2").getValue() || 1;

    const hoje = new Date();
    const hojeDia = parseInt(Utilities.formatDate(hoje, tZ, "dd"), 10);
    const hojeMes = parseInt(Utilities.formatDate(hoje, tZ, "MM"), 10) - 1;
    const hojeAno = parseInt(Utilities.formatDate(hoje, tZ, "yyyy"), 10);
    const hojeLimpo = new Date(hojeAno, hojeMes, hojeDia).getTime();

    let inicioSemana = getInicioDaSemana(hoje);

    let totalMes = 0,
        totalSemana = 0,
        totalDia = 0;
    let rankingModelos = {},
        rankingClientesMes = {},
        rankingSemanal = {};
    let detalheClienteModelo = {};

    for (let i = 0; i < dados.length; i++) {
        let linha = dados[i];
        let modelo = linha[5] ? linha[5].toString().trim() : "";
        let vendedorDaLinha = linha[9] ? linha[9].toString().trim() : "";

        if (modelo === "" || vendedorDaLinha !== vendedorSelecionado) continue;

        let status = linha[3] ? linha[3].toString().trim().toLowerCase() : "";
        if (status !== "pago") continue;

        let infoData = obterComponentesData(linha[2], tZ);
        if (!infoData.valida) continue;

        let dataPgtoLimpa = new Date(infoData.ano, infoData.mes, infoData.dia).getTime();
        let qtd = parseInt(linha[6]) || 0;
        let cliente = linha[4] ? linha[4].toString().trim() : "Desconhecido";

        if (dataPgtoLimpa === hojeLimpo) totalDia += qtd;

        if (infoData.mes === hojeMes && infoData.ano === hojeAno) {
            totalMes += qtd;
            rankingModelos[modelo] = (rankingModelos[modelo] || 0) + qtd;
            rankingClientesMes[cliente] = (rankingClientesMes[cliente] || 0) + qtd;

            if (!detalheClienteModelo[cliente]) detalheClienteModelo[cliente] = {};
            detalheClienteModelo[cliente][modelo] = (detalheClienteModelo[cliente][modelo] || 0) + qtd;
        }

        if (dataPgtoLimpa >= inicioSemana && dataPgtoLimpa <= hojeLimpo) {
            totalSemana += qtd;
            rankingSemanal[cliente] = (rankingSemanal[cliente] || 0) + qtd;
        }
    }

    dados = null;

    const ordenar = (obj, limite = 15) =>
        Object.entries(obj)
            .sort((a, b) => b[1] - a[1])
            .slice(0, limite);

    let listaDetalhada = [];
    for (let c in detalheClienteModelo) {
        for (let m in detalheClienteModelo[c]) {
            listaDetalhada.push([c, m, detalheClienteModelo[c][m]]);
        }
    }
    listaDetalhada.sort((a, b) => b[2] - a[2]);

    const payload = {
        metaMes: metaReal,
        projecao: metaHipotetica,
        vendasDia: totalDia,
        progressoDia: ((totalDia / (metaReal / 30)) * 100).toFixed(1),
        vendasSemana: totalSemana,
        progressoSemana: ((totalSemana / (metaReal / 4)) * 100).toFixed(1),
        vendasMes: totalMes,
        progressoMes: ((totalMes / metaReal) * 100).toFixed(1),
        progressoProj: ((totalMes / metaHipotetica) * 100).toFixed(1),
        topModelos: ordenar(rankingModelos, 15),
        topCompradoresMes: ordenar(rankingClientesMes, 15),
        topCompradoresSemana: ordenar(rankingSemanal, 15),
        detalhe: listaDetalhada.slice(0, 15)
    };

    const htmlTemplate = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>Dashboard Executivo</title>
      <script src="https://cdn.tailwindcss.com"></script>
      <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
      <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
      <style>
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #1e293b; }
        ::-webkit-scrollbar-thumb { background: #475569; border-radius: 4px; }
      </style>
    </head>
    <body class="bg-slate-900 text-slate-100 p-6 min-h-screen">
      <div class="max-w-7xl mx-auto space-y-6">
        <div class="flex justify-between items-center bg-slate-800 p-4 rounded-xl border border-slate-700 shadow-lg">
          <div>
            <h1 class="text-2xl font-bold text-slate-100 flex items-center gap-3">
              <i class="fa-solid fa-chart-line text-blue-400"></i>
              DESEMPENHO: <span class="text-blue-400">${vendedorSelecionado.toUpperCase()}</span>
            </h1>
            <p class="text-xs text-slate-400 mt-1">DEPARTAMENTO DE VENDAS SÃO BERNARDO DO CAMPO - SP</p>
          </div>
          <button onclick="google.script.host.close()" class="bg-slate-700 hover:bg-rose-600 text-white font-medium px-4 py-2 rounded-lg text-sm transition-colors shadow flex items-center gap-2 border border-slate-600">
            <i class="fa-solid fa-xmark"></i> Fechar
          </button>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-5 gap-4">
          <div class="bg-slate-800 p-4 rounded-xl border border-slate-700 shadow text-center">
            <p class="text-xs text-slate-400 font-medium uppercase"><i class="fa-solid fa-bolt text-amber-400 mr-1"></i> Dia</p>
            <p class="text-3xl font-black text-amber-400 my-1">${payload.vendasDia}</p>
            <div class="w-full bg-slate-700 h-2 rounded-full overflow-hidden mt-2">
              <div class="bg-amber-400 h-full" style="width: ${Math.min(payload.progressoDia, 100)}%"></div>
            </div>
            <p class="text-[10px] text-right text-slate-400 mt-1">${payload.progressoDia}%</p>
          </div>
          <div class="bg-slate-800 p-4 rounded-xl border border-slate-700 shadow text-center">
            <p class="text-xs text-slate-400 font-medium uppercase"><i class="fa-solid fa-calendar-week text-emerald-400 mr-1"></i> Semana</p>
            <p class="text-3xl font-black text-emerald-400 my-1">${payload.vendasSemana}</p>
            <div class="w-full bg-slate-700 h-2 rounded-full overflow-hidden mt-2">
              <div class="bg-emerald-400 h-full" style="width: ${Math.min(payload.progressoSemana, 100)}%"></div>
            </div>
            <p class="text-[10px] text-right text-slate-400 mt-1">${payload.progressoSemana}%</p>
          </div>
          <div class="bg-slate-800 p-4 rounded-xl border border-slate-700 shadow text-center">
            <p class="text-xs text-slate-400 font-medium uppercase"><i class="fa-solid fa-calendar-check text-blue-400 mr-1"></i> Mês</p>
            <p class="text-3xl font-black text-blue-400 my-1">${payload.vendasMes}</p>
            <div class="w-full bg-slate-700 h-2 rounded-full overflow-hidden mt-2">
              <div class="bg-blue-400 h-full" style="width: ${Math.min(payload.progressoMes, 100)}%"></div>
            </div>
            <p class="text-[10px] text-right text-slate-400 mt-1">${payload.progressoMes}%</p>
          </div>
          <div class="bg-slate-800 p-4 rounded-xl border border-slate-700 shadow text-center">
            <p class="text-xs text-slate-400 font-medium uppercase"><i class="fa-solid fa-bullseye text-slate-200 mr-1"></i> Meta</p>
            <p class="text-3xl font-black text-slate-200 mt-1">${payload.metaMes}</p>
          </div>
          <div class="bg-slate-800 p-4 rounded-xl border border-slate-700 shadow text-center">
            <p class="text-xs text-slate-400 font-medium uppercase"><i class="fa-solid fa-rocket text-purple-400 mr-1"></i> Projeção</p>
            <p class="text-3xl font-black text-purple-400 my-1">${payload.projecao}</p>
          </div>
        </div>

        <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div class="bg-slate-800 p-4 rounded-xl border border-slate-700 shadow flex flex-col h-[320px] col-span-1">
            <h2 class="text-sm font-bold text-pink-400 mb-3 border-b border-slate-700 pb-2 uppercase tracking-wider flex items-center gap-2">
              <i class="fa-solid fa-chart-pie"></i> Participação por Modelo
            </h2>
            <div class="flex-1 relative w-full h-full flex justify-center items-center pb-2">
              ${payload.topModelos.length > 0 ? '<canvas id="graficoModelos"></canvas>' : '<div class="text-slate-500 text-center italic">Sem dados</div>'}
            </div>
          </div>
          <div class="bg-slate-800 p-4 rounded-xl border border-slate-700 shadow flex flex-col h-[320px] col-span-2">
            <h2 class="text-sm font-bold text-purple-400 mb-3 border-b border-slate-700 pb-2 uppercase tracking-wider flex items-center gap-2">
              <i class="fa-solid fa-magnifying-glass-chart"></i> Últimas Operações (Top 15)
            </h2>
            <div class="overflow-y-auto flex-1 pr-2">
              <table class="w-full text-xs text-left">
                <tr class="text-slate-400 border-b border-slate-700 sticky top-0 bg-slate-800 z-10">
                  <th class="pb-2">Cliente</th><th class="pb-2">Modelo</th><th class="pb-2 text-right">Qtd</th>
                </tr>
                ${payload.detalhe
                    .map(
                        ([cli, mod, qtd]) => `
                  <tr class="border-b border-slate-700/50 hover:bg-slate-700/50">
                    <td class="py-2.5 text-slate-300">${cli}</td><td class="py-2.5 text-slate-400">${mod}</td><td class="py-2.5 text-right font-bold text-purple-400">${qtd}</td>
                  </tr>`
                    )
                    .join("")}
              </table>
            </div>
          </div>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div class="bg-slate-800 p-4 rounded-xl border border-slate-700 shadow flex flex-col h-[350px]">
            <h2 class="text-sm font-bold text-orange-400 mb-3 border-b border-slate-700 pb-2 uppercase"><i class="fa-solid fa-trophy mr-1"></i> Top Modelos Mais Vendidos</h2>
            <div class="overflow-y-auto flex-1 pr-2">
              <table class="w-full text-xs text-left">
                ${payload.topModelos.map(([mod, qtd]) => `<tr class="border-b border-slate-700/50"><td class="py-2.5 text-slate-300">${mod}</td><td class="py-2.5 text-right font-bold text-orange-400">${qtd}</td></tr>`).join("")}
              </table>
            </div>
          </div>
          <div class="bg-slate-800 p-4 rounded-xl border border-slate-700 shadow flex flex-col h-[350px]">
            <h2 class="text-sm font-bold text-blue-400 mb-3 border-b border-slate-700 pb-2 uppercase"><i class="fa-solid fa-user-tie mr-1"></i> Top Clientes do Mês</h2>
            <div class="overflow-y-auto flex-1 pr-2">
              <table class="w-full text-xs text-left">
                ${payload.topCompradoresMes.map(([cli, qtd]) => `<tr class="border-b border-slate-700/50"><td class="py-2.5 text-slate-300">${cli}</td><td class="py-2.5 text-right font-bold text-blue-400">${qtd}</td></tr>`).join("")}
              </table>
            </div>
          </div>
          <div class="bg-slate-800 p-4 rounded-xl border border-slate-700 shadow flex flex-col h-[350px]">
            <h2 class="text-sm font-bold text-emerald-400 mb-3 border-b border-slate-700 pb-2 uppercase"><i class="fa-solid fa-calendar-day mr-1"></i> Top Clientes da Semana</h2>
            <div class="overflow-y-auto flex-1 pr-2">
              <table class="w-full text-xs text-left">
                ${payload.topCompradoresSemana.map(([cli, qtd]) => `<tr class="border-b border-slate-700/50"><td class="py-2.5 text-slate-300">${cli}</td><td class="py-2.5 text-right font-bold text-emerald-400">${qtd}</td></tr>`).join("")}
              </table>
            </div>
          </div>
        </div>
      </div>
      <script>
        const cData = ${JSON.stringify(payload.topModelos)};
        if(cData.length > 0){
          new Chart(document.getElementById('graficoModelos').getContext('2d'), {
            type: 'doughnut',
            data: { labels: cData.map(i=>i[0]), datasets: [{ data: cData.map(i=>i[1]), backgroundColor: ['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#ec4899','#14b8a6'], borderWidth: 0 }] },
            options: { 
              responsive: true, 
              maintainAspectRatio: false, 
              layout: {
                padding: 15 // AQUI ESTÁ A CORREÇÃO DA MARGEM!
              },
              plugins: { 
                legend: { position: 'right', labels: { color: '#94a3b8', font: {size: 10} } } 
              }, 
              cutout: '65%' 
            }
          });
        }
      </script>
    </body>
    </html>
  `;

    SpreadsheetApp.getUi().showModalDialog(
        HtmlService.createHtmlOutput(htmlTemplate).setWidth(1400).setHeight(950),
        `🌐 Gerencial: ${vendedorSelecionado}`
    );
}

// ====================================================================
// DASHBOARD GERAL INTERNA DA PLANILHA (OTIMIZADA PARA MEMÓRIA)
// ====================================================================
function atualizarDashboardCompleto() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abaDados = ss.getSheetByName("Banco de Dados");
    const abaDash = ss.getSheetByName("DashBoard");
    const tZ = ss.getSpreadsheetTimeZone();

    // Limpa as tabelas antes de atualizar para evitar dados fantasmas
    abaDash.getRange("A20:B35").clearContent();
    abaDash.getRange("D20:E35").clearContent();
    abaDash.getRange("G20:H35").clearContent();
    abaDash.getRange("J20:L35").clearContent();

    const ultimaLinha = abaDados.getLastRow();
    if (ultimaLinha < 2) return;

    const maxLinhas = 5000;
    const linhasParaLer = Math.min(ultimaLinha - 1, maxLinhas);
    const linhaInicial = ultimaLinha - linhasParaLer + 1;
    let dados = abaDados.getRange(linhaInicial, 1, linhasParaLer, 10).getValues();

    const metaReal = abaDash.getRange("B2").getValue() || 1;
    const metaHipotetica = abaDash.getRange("C2").getValue() || 1;

    const hoje = new Date();
    const hojeLimpo = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate()).getTime();
    let inicioSemana = getInicioDaSemana(hoje);

    let totalMes = 0,
        totalSemana = 0,
        totalDia = 0;
    let rModelos = {},
        rClientes = {},
        rSemanal = {},
        det = {};

    for (let i = 0; i < dados.length; i++) {
        let linha = dados[i];
        let modelo = linha[5] ? linha[5].toString().trim() : "";
        let status = linha[3] ? linha[3].toString().trim().toLowerCase() : "";

        if (modelo === "" || status !== "pago") continue;

        let infoData = obterComponentesData(linha[2], tZ);
        if (!infoData.valida) continue;

        let dataPgto = new Date(infoData.ano, infoData.mes, infoData.dia).getTime();
        let qtd = parseInt(linha[6]) || 0;
        let cliente = linha[4] ? linha[4].toString().trim() : "Desconhecido";

        if (dataPgto === hojeLimpo) totalDia += qtd;
        if (infoData.mes === hoje.getMonth() && infoData.ano === hoje.getFullYear()) {
            totalMes += qtd;
            rModelos[modelo] = (rModelos[modelo] || 0) + qtd;
            rClientes[cliente] = (rClientes[cliente] || 0) + qtd;
            if (!det[cliente]) det[cliente] = {};
            det[cliente][modelo] = (det[cliente][modelo] || 0) + qtd;
        }
        if (dataPgto >= inicioSemana && dataPgto <= hojeLimpo) {
            totalSemana += qtd;
            rSemanal[cliente] = (rSemanal[cliente] || 0) + qtd;
        }
    }

    dados = null; // Libera memória imediatamente

    // Escreve os KPIs agregados
    abaDash.getRange("E2:F5").setValues([
        [totalDia, totalDia / (metaReal / 30)],
        [totalSemana, totalSemana / (metaReal / 4)],
        [totalMes, totalMes / metaReal],
        ["", totalMes / metaHipotetica]
    ]);

    const ordenar = (obj) =>
        Object.entries(obj)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 15);
    let lModelos = ordenar(rModelos),
        lClientes = ordenar(rClientes),
        lSemanal = ordenar(rSemanal);

    // Garante preenchimento correto prevenindo matrizes vazias ou assíncronas
    abaDash.getRange(20, 1, 15, 2).setValues(Array.from({ length: 15 }, (_, i) => lModelos[i] || ["", ""]));
    abaDash.getRange(20, 4, 15, 2).setValues(Array.from({ length: 15 }, (_, i) => lClientes[i] || ["", ""]));
    abaDash.getRange(20, 7, 15, 2).setValues(Array.from({ length: 15 }, (_, i) => lSemanal[i] || ["", ""]));

    let lDet = [];
    for (let c in det) for (let m in det[c]) lDet.push([c, m, det[c][m]]);
    lDet.sort((a, b) => b[2] - a[2]);
    if (lDet.length > 0) {
        abaDash
            .getRange(20, 10, Math.min(lDet.length, 15), 3)
            .setValues(Array.from({ length: 15 }, (_, i) => lDet[i] || ["", "", ""]));
    }

    ss.toast("Dashboard Geral atualizada com sucesso!", "📊 Sistema");
}
