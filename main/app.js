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

document.getElementById("funcionalidade").addEventListener("change", (e) => {
const funcionalidade = e.target.value;
const instrucoesDiv = document.getElementById("instrucoes");
const instrucoesText = document.getElementById("instrucoes-texto");

let texto = "";
if (funcionalidade === "relatorio") {
 texto = "Usado para criar a planilha dos procedimentos realizados de guias manuais de todos os laboratórios.<br><br>~> No site do SIGS (https://franca.sp.gov.br/sigs/) baixe um único arquivo com todos os laboratórios, respeitando as datas de realização do mês anterior e processe usando esta funcionalidade. ";
} else if (funcionalidade === "glosas") {
 texto = "Usado para conferir os procedimentos de guias manuais de todos os laboratorios.<br><br>~> Usando o mesmo arquivo baixado para a funcionalidade anterior, processe e use para conferência. ";
} else if (funcionalidade === "bdp") {
 texto = "Usado para lançar os procedimentos no BDP do SIA e fazer os débitos no faturamento de cada laboratório.<br><br>~> Use o arquivo de glosas (já conferido) com todos os registros glosados para processamento do montante de débitos.";
}

if (texto) {
 instrucoesText.innerHTML = texto;
 instrucoesDiv.style.display = "block";
} else {
 instrucoesDiv.style.display = "none";
}
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

 sheet.getColumn(1).width = 115;
 sheet.getColumn(2).width = 10;
 sheet.getColumn(3).width = 15;
 sheet.getColumn(4).width = 15;
 sheet.getColumn(5).width = 10;
 sheet.getColumn(6).width = 40;
 sheet.getColumn(7).width = 10;
 sheet.getColumn(8).width = 15;
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

 sheet.getColumn(1).width = 115;
 sheet.getColumn(2).width = 10;
 sheet.getColumn(3).width = 15;
 sheet.getColumn(4).width = 15;
 sheet.getColumn(5).width = 10;
 sheet.getColumn(6).width = 40;
 sheet.getColumn(7).width = 10;
 sheet.getColumn(8).width = 15;
 sheet.getColumn(9).width = 40;
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
function parseCurrencyBR(value) {
if (value == null) return 0;
if (typeof value === "number") return value;
let s = String(value).trim();
if (!s) return 0;
s = s.replace(/[^0-9.,-]/g, "");
if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
const n = parseFloat(s);
return isNaN(n) ? 0 : n;
}

function detectarCabecalhoGlosas(json) {
for (let i = 0; i < Math.min(json.length, 25); i++) {
 const row = json[i] || [];
 const lower = row.map(v => (v == null ? "" : String(v)).toLowerCase());
 const idxProc = lower.findIndex(t => t.includes("desc") && t.includes("proced"));
 const idxPreco = lower.findIndex(t => t.includes("r$") || t.includes("proc.") || t.includes("preço") || t.includes("preco"));
 if (idxProc !== -1 && idxPreco !== -1) {
 return {
  headerRow: i,
  cols: { procedimento: idxProc, preco: idxPreco }
 };
 }
}
return { headerRow: 1, cols: { procedimento: 0, preco: 1 } };
}

async function consolidacaoDeBdp() {
const wbOut = new ExcelJS.Workbook();
for (const sheetName of workbook.SheetNames) {
 const ws = workbook.Sheets[sheetName];
 const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
 if (!json || !json.length) continue;
 const { headerRow, cols } = detectarCabecalhoGlosas(json);
 const start = headerRow + 1;
 const mapa = new Map();
 for (let r = start; r < json.length; r++) {
 const row = json[r] || [];
 const proc = row[cols.procedimento];
 if (!proc) continue;
 const precoRaw = cols.preco >= 0 ? row[cols.preco] : 0;
 const valorUnit = parseCurrencyBR(precoRaw);
 const chave = `${proc}||${valorUnit.toFixed(2)}`;
 const atual = mapa.get(chave) || { procedimento: proc, valorUnitario: valorUnit, quantidade: 0 };
 atual.quantidade += 1;
 mapa.set(chave, atual);
 }
 let linhas = Array.from(mapa.values()).map(x => [
 x.procedimento, x.valorUnitario, x.quantidade, x.valorUnitario * x.quantidade
 ]);
 linhas.sort((a, b) => {
 if (b[2] !== a[2]) return b[2] - a[2];
 return String(a[0]).localeCompare(String(b[0]));
 });
 const nomeLab = resolverNomeLaboratorio(sheetName);
 const sheet = wbOut.addWorksheet(nomeLab);
 sheet.mergeCells(1, 1, 1, 4);
 const titulo = sheet.getCell(1, 1);
 titulo.value = `BDP - ${nomeLab}`;
 titulo.font = { name: "Arial", size: 12, bold: true, underline: true };
 titulo.alignment = { horizontal: "center" };
 const headerExcelRow = sheet.addRow(["Procedimento", "Valor unitário", "Quantidade", "Valor total"]);
 headerExcelRow.font = { bold: true, name: "Arial" };
 for (const l of linhas) sheet.addRow(l);
 const totalQtd = linhas.reduce((s, r) => s + (r[2] || 0), 0);
 const totalVal = linhas.reduce((s, r) => s + (r[3] || 0), 0);
 const totalRow = sheet.addRow(["", "Total", totalQtd, totalVal]);
 totalRow.eachCell(cell => {
 cell.font = { bold: true };
 cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };
 cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
 });
 totalRow.getCell(2).numFmt = '"R$"#,##0.00';
 totalRow.getCell(3).numFmt = '0';
 totalRow.getCell(4).numFmt = '"R$"#,##0.00';
 sheet.eachRow({ includeEmpty: false }, (row) => {
 row.eachCell({ includeEmpty: false }, (cell) => {
  cell.font = cell.font || { name: "Calibri", size: 12 };
  cell.border = { top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" } };
 });
 });
 
 sheet.getColumn(1).width = 115;
 sheet.getColumn(2).width = 10;
 sheet.getColumn(3).width = 10;
 sheet.getColumn(4).width = 10;
}
const buffer = await wbOut.xlsx.writeBuffer();
const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
const url = URL.createObjectURL(blob);
const a = document.createElement("a");
a.href = url;
a.download = "BDP (mês e ano).xlsx";
a.click();
URL.revokeObjectURL(url);
}