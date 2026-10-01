// Google Apps Script — grava no Google Sheets o que for alterado no site.
// Cole em: planilha > Extensões > Apps Script. Depois: Implantar > Nova implantação > App da Web
// Executar como: Eu | Quem tem acesso: Qualquer pessoa. Copie a URL /exec para SCRIPT_URL no setores.html.
const GID = 1534942531;
const HEADERS = ["Setor","EUV","EDOC","Tipo","Bairro","Status","Telefone","Tem whatsapp?","Telefone 2","Tem whatsapp? 2","E-mail","Endereço"];
const KEYS    = ["setor","euv","edoc","tipo","bairro","ok","tel","w1","tel2","w2","email","end"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const m = JSON.parse(e.postData.contents);
    const sh = SpreadsheetApp.getActive().getSheets().filter(s => s.getSheetId() === GID)[0];
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
    const col = k => head.indexOf(HEADERS[KEYS.indexOf(k)].toLowerCase()) + 1;
    const cs = col("setor");
    const find = () => {
      const n = sh.getLastRow() - 1;
      if (n < 1) return 0;
      const v = sh.getRange(2, cs, n, 1).getDisplayValues().flat();
      const i = v.findIndex(x => String(x).trim() === String(m.setor).trim());
      return i < 0 ? 0 : i + 2;
    };
    const val = (k, v) => k === "ok" ? (v === true || v === "true") : v;

    if (m.op === "set") {
      const r = find(), c = col(m.k);
      if (r && c) sh.getRange(r, c).setValue(val(m.k, m.v));
    } else if (m.op === "add") {
      if (!find()) {
        const row = head.map(() => "");
        KEYS.forEach(k => { const c = col(k); if (c) row[c - 1] = val(k, m.row[k]); });
        sh.appendRow(row);
      }
    } else if (m.op === "del") {
      const r = find();
      if (r) sh.deleteRow(r);
    }
    SpreadsheetApp.flush();
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}
