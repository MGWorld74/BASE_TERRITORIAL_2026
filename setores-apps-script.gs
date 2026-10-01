// Cole em: planilha > Extensões > Apps Script. Depois: Implantar > Nova implantação > App da Web
// (Executar como: Eu | Quem pode acessar: Qualquer pessoa). Copie a URL para SCRIPT_URL no setores.html.
const COLS = ["setor","euv","edoc","tipo","bairro","ok","tel","w1","tel2","w2","email","end"];

function doPost(e) {
  const m = JSON.parse(e.postData.contents);
  const sh = SpreadsheetApp.getActive().getSheets().find(s => String(s.getSheetId()) === String(m.gid));
  const last = Math.max(sh.getLastRow(), 2);
  const ids = sh.getRange(1, 1, last, 1).getValues().map(r => String(r[0]));
  const row = ids.indexOf(String(m.setor)) + 1; // 0 se não achou

  if (m.op === "set" && row > 1) {
    sh.getRange(row, COLS.indexOf(m.k) + 1).setValue(m.v);
  } else if (m.op === "del" && row > 1) {
    sh.deleteRow(row);
  } else if (m.op === "add" && row < 1) {
    sh.appendRow(COLS.map(c => m.row[c]));
  }
  return ContentService.createTextOutput("ok");
}
