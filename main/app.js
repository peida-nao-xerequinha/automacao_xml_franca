let workbook;

document.getElementById("file-input").addEventListener("change", function (e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    workbook = XLSX.read(data, { type: 'array' });
  };

  reader.readAsArrayBuffer(file);
});

document.getElementById("executar-btn").addEventListener("click", () => {
  const funcionalidade = document.getElementById("funcionalidade").value;

  if (!workbook) {
    alert("Por favor, selecione um arquivo primeiro.");
    return;
  }

  if (funcionalidade === "relatorio") {
    executarRelatorio();
  } else if (funcionalidade === "glosas") {
    executarPlanilhaDeGlosas();
  } else if (funcionalidade === "bdp") {
    consolidacaoDeBdp();
  } else {
    alert("Funcionalidade ainda não implementada.");
  }
});

// === RESOLVER NOME DE LABORATÓRIO ===
function normalizarTexto(txt) {
  return (txt || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase();
}

function resolverNomeLaboratorio(labRaw) {
  const nome = normalizarTexto(labRaw);

  if (nome.includes("SANTA CASA")) return "Santa Casa";
  if (nome.includes("BIOANALISE")) return "Laboratório Bioanálises";
  if (nome.includes("LABCENTER")) return "Labcenter";

  return labRaw.substring(0, 31); // fallback
}

// === AUTO AJUSTE DE COLUNAS ===
function autoFitColumns(sheet) {
  sheet.columns.forEach(col => {
    let maxLength = 10;
    col.eachCell({ includeEmpty: true }, (cell) => {
      const val = cell.value ? cell.value.toString() : "";
      maxLength = Math.max(maxLength, val.length + 2);
    });
    col.width = Math.min(Math.max(maxLength, 15), 40); // entre 15 e 40
  });
}

// === RELATÓRIO POR LABORATÓRIO ===
async function executarRelatorio() {
  const originalSheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(originalSheet, { header: 1 });
  const dados = json.slice(1);

  const colIndices = {
    laboratorio: 0, procedimento: 2, preco: 3,
    dataSolicitacao: 5, dataRealizacao: 6,
    matricula: 7, paciente: 8, crm: 9, envioXml: 11
  };

  const separados = {};
  for (const linha of dados) {
    const lab = linha[colIndices.laboratorio];
    if (!lab || lab === "Laboratório") continue;
    if (!separados[lab]) separados[lab] = [];

    const procedimento = (linha[colIndices.procedimento] || "").trim();
    let precoStr = linha[colIndices.preco]?.toString().replace('R$', '').trim() || "0";
    precoStr = precoStr.replace('.', '').replace(',', '.');
    const preco = parseFloat(precoStr) || 0;

    separados[lab].push([
      procedimento, preco,
      linha[colIndices.dataSolicitacao], linha[colIndices.dataRealizacao],
      linha[colIndices.matricula], linha[colIndices.paciente],
      linha[colIndices.crm], linha[colIndices.envioXml]
    ]);
  }

  const workbookExcelJS = new ExcelJS.Workbook();

  for (const lab in separados) {
    const linhas = separados[lab];
    linhas.sort((a, b) => (a[5] || "").localeCompare(b[5] || ""));
    const nomeLab = resolverNomeLaboratorio(lab);
    const sheet = workbookExcelJS.addWorksheet(nomeLab);

    sheet.mergeCells("A1:H1");
    const titulo = sheet.getCell("A1");
    titulo.value = `XML Guias Manuais - ${nomeLab}`;
    titulo.font = { name: "Arial", size: 12, bold: true, underline: true };
    titulo.alignment = { horizontal: "center" };

    const cabecalho = [
      "Desc. Procedimento","R$ Proc.","Data Solicitação","Data Realização",
      "MCV","Nome Paciente","CRM","XML"
    ];
    sheet.addRow(cabecalho).font = { bold: true, name: "Arial" };

    for (const linha of linhas) sheet.addRow(linha);

    sheet.getColumn(2).numFmt = '"R$"#,##0.00';

    sheet.eachRow(row => row.eachCell(cell => {
      cell.font = cell.font || { name: "Calibri", size: 12 };
      cell.border = { top:{style:"thin"},left:{style:"thin"},bottom:{style:"thin"},right:{style:"thin"} };
    }));

    autoFitColumns(sheet);
  }

  const buffer = await workbookExcelJS.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "XML (mês e ano).xlsx";
  a.click();
  URL.revokeObjectURL(url);
}

// === PLANILHA DE GLOSAS ===
async function executarPlanilhaDeGlosas() {
  const originalSheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(originalSheet, { header: 1 });
  const dados = json.slice(1);

  const colIndices = {
    laboratorio: 0, procedimento: 2, preco: 3,
    dataSolicitacao: 5, dataRealizacao: 6,
    matricula: 7, paciente: 8, crm: 9, envioXml: 11
  };

  const separados = {};
  for (const linha of dados) {
    const lab = linha[colIndices.laboratorio];
    if (!lab || lab === "Laboratório") continue;
    if (!separados[lab]) separados[lab] = [];

    const procedimento = (linha[colIndices.procedimento] || "").trim();
    let precoStr = linha[colIndices.preco]?.toString().replace('R$', '').trim() || "0";
    precoStr = precoStr.replace('.', '').replace(',', '.');
    const preco = parseFloat(precoStr) || 0;

    separados[lab].push([
      procedimento, preco,
      linha[colIndices.dataSolicitacao], linha[colIndices.dataRealizacao],
      linha[colIndices.matricula], linha[colIndices.paciente],
      linha[colIndices.crm], linha[colIndices.envioXml],
      "" // coluna glosa
    ]);
  }

  const workbookExcelJS = new ExcelJS.Workbook();

  for (const lab in separados) {
    const linhas = separados[lab];
    linhas.sort((a, b) => (a[5] || "").localeCompare(b[5] || ""));
    const nomeLab = resolverNomeLaboratorio(lab);
    const sheet = workbookExcelJS.addWorksheet(nomeLab);

    sheet.mergeCells("A1:I1");
    const titulo = sheet.getCell("A1");
    titulo.value = `Glosas Guias Manuais - ${nomeLab}`;
    titulo.font = { name: "Arial", size: 12, bold: true, underline: true };
    titulo.alignment = { horizontal: "center" };

    const cabecalho = [
      "Desc. Procedimento","R$ Proc.","Data Solicitação","Data Realização",
      "MCV","Nome Paciente","CRM","XML","Glosa"
    ];
    sheet.addRow(cabecalho).font = { bold: true, name: "Arial" };

    for (const linha of linhas) sheet.addRow(linha);

    sheet.getColumn(2).numFmt = '"R$"#,##0.00';

    sheet.eachRow(row => row.eachCell(cell => {
      cell.font = cell.font || { name: "Calibri", size: 12 };
      cell.border = { top:{style:"thin"},left:{style:"thin"},bottom:{style:"thin"},right:{style:"thin"} };
    }));

    autoFitColumns(sheet);
  }

  const buffer = await workbookExcelJS.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "Glosas (mês e ano).xlsx";
  a.click();
  URL.revokeObjectURL(url);
}

// === CONSOLIDAÇÃO DE BDP ===
// --- helpers novos ---
function parseCurrencyBR(value) {
  if (value == null) return 0;
  if (typeof value === "number") return value;
  let s = String(value).trim();
  if (!s) return 0;
  // remove tudo que não é dígito, ponto, vírgula ou sinal
  s = s.replace(/[^0-9.,-]/g, "");
  // se tiver vírgula, tratamos como decimal BR: remove pontos (milhar) e troca vírgula por ponto
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function detectarCabecalhoGlosas(json) {
  // tenta achar a linha que contenha "Desc." e "R$" (ou similar)
  for (let i = 0; i < Math.min(json.length, 25); i++) {
    const row = json[i] || [];
    const lower = row.map(v => (v == null ? "" : String(v)).toLowerCase());
    const idxProc = lower.findIndex(t => t.includes("desc") && t.includes("proced"));
    const idxPreco = lower.findIndex(t => t.includes("r$") || t.includes("proc.") || t.includes("preço") || t.includes("preco"));
    const idxQtd = lower.findIndex(t => t.includes("quant")); // não é obrigatório
    if (idxProc !== -1 && idxPreco !== -1) {
      // mapeia outras colunas úteis se existir
      const idxXML = lower.findIndex(t => t.trim() === "xml" || t.includes("xml"));
      const idxGlosa = lower.findIndex(t => t.includes("glosa"));
      return {
        headerRow: i,
        cols: { procedimento: idxProc, preco: idxPreco, xml: idxXML, glosa: idxGlosa }
      };
    }
  }
  // fallback para layout padrão: A=proced, B=preço
  return { headerRow: 1, cols: { procedimento: 0, preco: 1, xml: -1, glosa: -1 } };
}

// usa sua função existente:
function resolverNomeLaboratorio(labRaw) {
  const nome = (labRaw || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase();

  if (nome.includes("SANTA CASA")) return "Santa Casa";
  if (nome.includes("BIOANALISE")) return "Laboratório Bioanálises";
  if (nome.includes("LABCENTER")) return "Labcenter";

  // se já veio bonitinho do arquivo de glosas, mantemos
  const trimmed = (labRaw || "").toString().trim();
  return trimmed ? trimmed.substring(0, 31) : "Planilha";
}

function autoFitColumns(sheet) {
  sheet.columns.forEach(col => {
    let maxLength = 10;
    col.eachCell({ includeEmpty: true }, (cell) => {
      let v = cell.value;
      if (v == null) v = "";
      else if (typeof v === "number") v = v.toFixed(2); // para moeda ficar “contável”
      else v = String(v);
      maxLength = Math.max(maxLength, v.length + 2);
    });
    col.width = Math.min(Math.max(maxLength, 15), 60);
  });
}

// --- substitua sua função por esta ---
async function consolidacaoDeBdp() {
  const wbOut = new ExcelJS.Workbook();

  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { header: 1 });

    if (!json || !json.length) continue;

    // detecta o cabeçalho real
    const { headerRow, cols } = detectarCabecalhoGlosas(json);
    const start = headerRow + 1;

    const mapa = new Map();
    for (let r = start; r < json.length; r++) {
      const row = json[r] || [];
      const proc = row[cols.procedimento];
      if (!proc) continue; // sem procedimento, ignora linha
      const precoRaw = cols.preco >= 0 ? row[cols.preco] : 0;
      const valorUnit = parseCurrencyBR(precoRaw);

      // chave = procedimento + preço (para não misturar se o preço variar)
      const chave = `${proc}||${valorUnit.toFixed(2)}`;
      const atual = mapa.get(chave) || { procedimento: proc, valorUnitario: valorUnit, quantidade: 0 };
      atual.quantidade += 1;
      mapa.set(chave, atual);
    }

    // monta linhas
    let linhas = Array.from(mapa.values()).map(x => [
      x.procedimento,
      x.valorUnitario,
      x.quantidade,
      x.valorUnitario * x.quantidade
    ]);

    // ordena por quantidade desc, desempate por procedimento asc
    linhas.sort((a, b) => {
      if (b[2] !== a[2]) return b[2] - a[2];
      return String(a[0]).localeCompare(String(b[0]));
    });

    // nome “bonito” da aba
    const nomeLab = resolverNomeLaboratorio(sheetName);
    const sheet = wbOut.addWorksheet(nomeLab);

    // título
    sheet.mergeCells(1, 1, 1, 4);
    const titulo = sheet.getCell(1, 1);
    titulo.value = `BDP - ${nomeLab}`;
    titulo.font = { name: "Arial", size: 12, bold: true, underline: true };
    titulo.alignment = { horizontal: "center" };

    // cabeçalho
    const Rowheader = sheet.addRow(["Procedimento", "Valor unitário", "Quantidade", "Valor total"]);
    headerRow.font = { bold: true, name: "Arial" };

    // dados
    for (const l of linhas) sheet.addRow(l);

    // totais
    const totalQtd = linhas.reduce((s, r) => s + (r[2] || 0), 0);
    const totalVal = linhas.reduce((s, r) => s + (r[3] || 0), 0);
    const totalRow = sheet.addRow(["", "Total", totalQtd, totalVal]);

    // moeda
    sheet.getColumn(2).numFmt = '"R$"#,##0.00';
    sheet.getColumn(4).numFmt = '"R$"#,##0.00';

    // estilo borda + fonte
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        cell.font = cell.font || { name: "Calibri", size: 12 };
        cell.border = { top:{style:"thin"}, left:{style:"thin"}, bottom:{style:"thin"}, right:{style:"thin"} };
      });
    });

    // destaca SOMENTE B/C/D no total (A fica sem destaque)
    [2, 3, 4].forEach(ci => {
      const c = totalRow.getCell(ci);
      c.font = { ...(c.font || {}), bold: true };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };
    });

    autoFitColumns(sheet);
  }

  const buffer = await wbOut.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "Consolidacao_de_BDP.xlsx";
  a.click();
  URL.revokeObjectURL(url);
}
