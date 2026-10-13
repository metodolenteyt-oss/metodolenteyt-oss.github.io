/**
 * Método Lente — CRM dentro da planilha (grátis).
 * Crie um 2º arquivo no mesmo Apps Script (+ → Script → "CRM"), cole este código, salve,
 * recarregue a planilha e use o menu "Método Lente → Configurar CRM (1ª vez)".
 * A aba "Leads" continua sendo o registro bruto do site; nada aqui altera o formulário.
 */
const CRM_ABA = "CRM";
const PAINEL_ABA = "Painel";
const ETAPAS = ["Novo", "Contatado", "Interessado", "Negociando", "Comprou", "Perdido"];
const ETAPA_COR = { Novo: "#f1f1f1", Contatado: "#dcf6fd", Interessado: "#ece2ff",
                    Negociando: "#fae1f0", Comprou: "#c9f5d3", Perdido: "#f6caca" };
const CRM_COLS = ["Nome", "E-mail", "WhatsApp", "Chamar", "Aulas baixadas", "Downloads", "Jornada",
                  "1º contato", "Último contato", "Temperatura", "Etapa", "Próxima ação", "Observações"];
const MANUAIS = 3; // as 3 últimas colunas são preenchidas por você e nunca são apagadas
const MSG_WHATS = "Oi, {nome}! Aqui é a Profe Lu, do Método Lente. Vi que você baixou o mapa mental — " +
                  "posso te contar como o curso pode te ajudar no vestibular?";

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Método Lente")
    .addItem("Atualizar CRM agora", "atualizarCRM")
    .addSeparator()
    .addItem("Configurar CRM (1ª vez)", "configurarCRM")
    .addToUi();
}

function configurarCRM() {
  atualizarCRM();
  montarPainel_();
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === "atualizarCRM")
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("atualizarCRM").timeBased().everyHours(1).create();
  SpreadsheetApp.getActive().toast("CRM e Painel prontos. Atualização automática a cada hora.", "Método Lente", 6);
}

function atualizarCRM() {
  const ss = SpreadsheetApp.getActive();
  const leads = ss.getSheetByName("Leads");
  const crm = ss.getSheetByName(CRM_ABA) || ss.insertSheet(CRM_ABA, 0);
  if (!leads || leads.getLastRow() < 2) return;

  // 1) junta as linhas da aba Leads por e-mail
  const dados = leads.getDataRange().getValues();
  const h = dados.shift().map(String);
  const c = (nome) => h.indexOf(nome);
  const porEmail = new Map();
  dados.forEach((r) => {
    const email = String(r[c("email")] || "").trim().toLowerCase();
    if (!email) return;
    const quando = r[c("data")] instanceof Date ? r[c("data")] : new Date(r[c("data")]);
    const aula = String(r[c("aula")]).replace(/\D/g, "").padStart(2, "0");
    let m = porEmail.get(email);
    if (!m) {
      m = { nome: r[c("nome")], email, fone: r[c("whatsapp")], aulas: new Set(), n: 0,
            passos: [], primeiro: quando, ultimo: quando, whatsOk: false };
      porEmail.set(email, m);
    }
    m.aulas.add(aula);
    m.n++;
    const o = String(r[c("origem")] || "");
    m.passos.push({ quando, txt: aula + " · " + (o === "qr" ? "QR" : o === "yt" ? "Descrição" : o || "direto") +
      " · " + Utilities.formatDate(quando, Session.getScriptTimeZone(), "dd/MM") });
    if (quando < m.primeiro) m.primeiro = quando;
    if (quando >= m.ultimo) { m.ultimo = quando; m.nome = r[c("nome")] || m.nome; m.fone = r[c("whatsapp")] || m.fone; }
    if (String(r[c("consent_whatsapp")]).toLowerCase() === "sim") m.whatsOk = true;
  });

  // 2) preserva o que você escreveu (Etapa, Próxima ação, Observações)
  const manual = new Map();
  if (crm.getLastRow() > 1) {
    crm.getRange(2, 1, crm.getLastRow() - 1, CRM_COLS.length).getValues()
      .forEach((r) => manual.set(String(r[1]).toLowerCase(), r.slice(-MANUAIS)));
  }

  // 3) monta as linhas
  const linhas = [...porEmail.values()].sort((a, b) => b.ultimo - a.ultimo).map((m) => {
    const fone = String(m.fone).replace(/\D/g, "");
    const primeiro = String(m.nome).split(" ")[0];
    const chamar = m.whatsOk && fone.length >= 12
      ? '=HYPERLINK("https://wa.me/' + fone + "?text=" + encodeURIComponent(MSG_WHATS.replace("{nome}", primeiro)) + '","💬 Chamar")'
      : "sem permissão";
    const temp = m.aulas.size >= 3 ? "🔥 Quente" : m.aulas.size === 2 ? "🌤 Morno" : "❄ Frio";
    const [etapa, prox, obs] = manual.get(m.email) || ["Novo", "", ""];
    return [m.nome, m.email, "+" + fone, chamar, [...m.aulas].sort().join(", "), m.n,
            m.passos.sort((a, b) => a.quando - b.quando).map((x) => x.txt).join("  →  "), m.primeiro, m.ultimo, temp, etapa || "Novo", prox, obs];
  });

  // 4) escreve e formata
  crm.clearContents();
  crm.getRange(1, 1, 1, CRM_COLS.length).setValues([CRM_COLS]);
  if (linhas.length) crm.getRange(2, 1, linhas.length, CRM_COLS.length).setValues(linhas);
  formatarCRM_(crm, Math.max(linhas.length, 1));
}

function formatarCRM_(sh, n) {
  const W = CRM_COLS.length;
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  sh.getRange(1, 1, 1, W).setBackground("#252525").setFontColor("#ffffff").setFontWeight("bold")
    .setVerticalAlignment("middle");
  sh.setRowHeight(1, 34);
  [190, 220, 140, 95, 120, 90, 320, 120, 120, 105, 120, 120, 260].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const corpo = sh.getRange(2, 1, Math.max(n, 200), W);
  corpo.setFontFamily("Arial").setFontSize(10).setVerticalAlignment("middle");
  sh.getRange(2, 8, Math.max(n, 200), 2).setNumberFormat("dd/mm/yyyy hh:mm");
  sh.getRange(2, 12, Math.max(n, 200), 1).setNumberFormat("dd/mm/yyyy")
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  sh.getRange(2, 11, Math.max(n, 200), 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ETAPAS, true).build())
    .setFontWeight("bold");
  sh.getRange(2, 4, Math.max(n, 200), 1).setHorizontalAlignment("center").setFontColor("#6400f8").setFontWeight("bold");

  // linha inteira colorida pela etapa + próxima ação vencida em rosa
  const regras = ETAPAS.map((e) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$K2="' + e + '"').setBackground(ETAPA_COR[e]).setRanges([corpo]).build());
  regras.unshift(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($L2<>"",$L2<=TODAY(),$K2<>"Comprou",$K2<>"Perdido")')
    .setBackground("#f72a9f").setFontColor("#ffffff").setRanges([sh.getRange(2, 12, Math.max(n, 200), 1)]).build());
  sh.setConditionalFormatRules(regras);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, n + 1, W).createFilter();
}

function montarPainel_() {
  const ss = SpreadsheetApp.getActive();
  const old = ss.getSheetByName(PAINEL_ABA);
  if (old) ss.deleteSheet(old);
  const p = ss.insertSheet(PAINEL_ABA, 0);
  p.setHiddenGridlines(true);
  p.setColumnWidths(1, 8, 130);
  p.getRange("A1:H60").setFontFamily("Arial").setBackground("#f4f4f4");

  p.getRange("A1").setValue("Método Lente · Painel de vendas").setFontSize(20).setFontWeight("bold").setFontColor("#252525");
  p.getRange("A2").setValue("Atualiza sozinho a cada hora · menu Método Lente → Atualizar CRM agora").setFontColor("#8a3e6b");

  // cartões de números
  const kpis = [
    ["Leads", "=COUNTA(CRM!B2:B)", "#6400f8"],
    ["Novos hoje", "=COUNTIF(CRM!H2:H,\">=\"&TODAY())", "#01c8f1"],
    ["Quentes 🔥", "=COUNTIF(CRM!J2:J,\"🔥*\")", "#f72a9f"],
    ["Compraram", "=COUNTIF(CRM!K2:K,\"Comprou\")", "#1e9e4a"],
    ["Conversão", "=IFERROR(COUNTIF(CRM!K2:K,\"Comprou\")/COUNTA(CRM!B2:B),0)", "#252525"],
    ["Com WhatsApp ok", "=COUNTIF(CRM!D2:D,\"<>sem permissão\")-COUNTBLANK(CRM!D2:D)", "#8a3e6b"],
  ];
  kpis.forEach(([rot, f, cor], i) => {
    const col = 1 + i;
    p.getRange(4, col).setValue(rot).setFontColor("#ffffff").setBackground(cor).setFontWeight("bold").setHorizontalAlignment("center");
    const v = p.getRange(5, col).setFormula(f).setFontSize(26).setFontWeight("bold").setFontColor(cor)
      .setBackground("#ffffff").setHorizontalAlignment("center");
    if (rot === "Conversão") v.setNumberFormat("0.0%");
  });
  p.setRowHeight(5, 54);

  // funil por etapa
  p.getRange("A8").setValue("Funil").setFontWeight("bold").setFontSize(13);
  p.getRange(9, 1, 1, 2).setValues([["Etapa", "Leads"]]).setFontWeight("bold");
  ETAPAS.forEach((e, i) => {
    p.getRange(10 + i, 1).setValue(e).setBackground(ETAPA_COR[e]);
    p.getRange(10 + i, 2).setFormula('=COUNTIF(CRM!K2:K,"' + e + '")').setBackground("#ffffff");
  });

  // por aula e por origem
  p.getRange("D8").setValue("Leads por aula").setFontWeight("bold").setFontSize(13);
  p.getRange("D9").setFormula('=QUERY(Leads!E2:E,"select E, count(E) where E is not null group by E label E \'Aula\', count(E) \'Downloads\'",0)');
  p.getRange("G8").setValue("Origem").setFontWeight("bold").setFontSize(13);
  p.getRange("G9").setFormula('=QUERY(Leads!F2:F,"select F, count(F) where F <> \'\' group by F label F \'Origem\', count(F) \'Downloads\'",0)');

  // agenda
  p.getRange("A18").setValue("Contatos dos próximos 3 dias").setFontWeight("bold").setFontSize(13);
  p.getRange("A19").setFormula('=IFERROR(SORT(FILTER({CRM!A2:A,CRM!B2:B,CRM!K2:K,CRM!L2:L},CRM!L2:L<>"",CRM!L2:L<=TODAY()+3,CRM!K2:K<>"Comprou",CRM!K2:K<>"Perdido"),4,TRUE),"Nenhum contato agendado")');
  p.getRange("D19:D60").setNumberFormat("dd/mm/yyyy");

  // gráficos
  p.insertChart(p.newChart().asPieChart().addRange(p.getRange("A9:B15")).setPosition(8, 9, 0, 0)
    .setOption("title", "Funil de vendas").setOption("pieHole", 0.45)
    .setOption("colors", ETAPAS.map((e) => ({ Novo: "#bdbdbd", Contatado: "#01c8f1", Interessado: "#6400f8",
      Negociando: "#f72a9f", Comprou: "#1e9e4a", Perdido: "#c0392b" }[e]))).build());
  p.insertChart(p.newChart().asColumnChart().addRange(p.getRange("D9:E30")).setPosition(26, 9, 0, 0)
    .setOption("title", "Downloads por aula").setOption("colors", ["#6400f8"]).setOption("legend", { position: "none" }).build());
}
