// ===================================================================== CRM
// Aba CRM: 1 linha por pessoa (chave = e-mail). Colunas A–J são preenchidas pelo script;
// L–N (Etapa, Próxima ação, Observações) são da equipe e NUNCA são sobrescritas.
const CRM = "CRM", PAINEL = "Painel", CONFIG = "Config";
const CRM_COLS = ["Nome", "E-mail", "WhatsApp", "Abrir WhatsApp", "Aceita WhatsApp", "Aulas baixadas",
                  "Nº de aulas", "Origem (1º contato)", "Primeiro contato", "Última atividade",
                  "Temperatura", "Etapa", "Próxima ação", "Observações"];
const ETAPAS = ["Novo", "Contatado", "Interessado", "Negociando", "Comprou", "Perdido"];
const MSG_PADRAO = "Oi, {nome}! Aqui é a Profe Lu, do Método Lente. Vi que você baixou o mapa mental " +
                   "da aula {aula}. Posso te contar como funciona o curso completo?";

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Método Lente")
    .addItem("Configurar / atualizar CRM", "configurarCRM")
    .addItem("Reconstruir CRM a partir de Leads", "reconstruirCRM")
    .addToUi();
}

function aula2_(v) {
  const s = String(v == null ? "" : v).replace(/\D/g, "");
  return s ? ("0" + s).slice(-2) : "";
}

function crmUpsert_(email, p, quando) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CRM);
  if (!sh) return; // CRM ainda não configurado
  const aula = aula2_(p.aula);
  const zap = String(p.whatsapp || "").replace(/[^\d+]/g, "");
  const aceita = p.consent_whatsapp === "sim";
  const found = sh.getRange("B2:B").createTextFinder(email).matchEntireCell(true).findNext();
  if (found) {
    const r = found.getRow();
    const v = sh.getRange(r, 1, 1, 10).getValues()[0];
    const aulas = String(v[5]).split(",").map((x) => x.trim()).filter(String);
    if (aula && aulas.indexOf(aula) < 0) aulas.push(aula);
    aulas.sort();
    if (p.nome) sh.getRange(r, 1).setValue(limpa_(p.nome));
    if (zap) sh.getRange(r, 3).setValue("'" + zap);
    if (aceita) sh.getRange(r, 5).setValue("sim");
    sh.getRange(r, 6, 1, 2).setValues([[aulas.join(", "), aulas.length]]);
    sh.getRange(r, 10).setValue(quando);
  } else {
    sh.appendRow([limpa_(p.nome), email, zap ? "'" + zap : "", "", aceita ? "sim" : "não", aula,
                  aula ? 1 : 0, limpa_(p.origem), quando, quando, "", "Novo", "", ""]);
    formulasLinha_(sh, sh.getLastRow());
  }
}

function formulasLinha_(sh, r) {
  // D: link do WhatsApp com a mensagem da aba Config (só para quem autorizou)
  sh.getRange(r, 4).setFormulaR1C1(
    '=IF(RC2="","",IF(RC5<>"sim","sem autorização",HYPERLINK("https://wa.me/"&REGEXREPLACE(RC3,"\\D","")&"?text="&' +
    'ENCODEURL(SUBSTITUTE(SUBSTITUTE(Config!R2C2,"{nome}",PROPER(REGEXEXTRACT(RC1&" ","^\\S+"))),"{aula}",' +
    'REGEXEXTRACT(RC6&"","[^,]*$"))),"Abrir conversa")))');
  // K: temperatura = nº de aulas + quão recente foi a última atividade
  sh.getRange(r, 11).setFormulaR1C1(
    '=IF(RC2="","",IF(OR(RC7>=3,AND(RC7>=2,TODAY()-RC10<=14),AND(RC12="Interessado",TODAY()-RC10<=30)),"Quente",' +
    'IF(TODAY()-RC10<=30,"Morno","Frio")))');
}

function configurarCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  aba_();
  // Config
  const cf = ss.getSheetByName(CONFIG) || ss.insertSheet(CONFIG);
  if (!cf.getRange("B2").getValue()) {
    cf.getRange("A1:B1").setValues([["Configuração", "Valor"]]).setFontWeight("bold").setBackground("#fae1f0");
    cf.getRange("A2:B2").setValues([["Mensagem de WhatsApp ({nome} e {aula} são trocados automaticamente)", MSG_PADRAO]]);
    cf.setColumnWidth(1, 330); cf.setColumnWidth(2, 620); cf.getRange("B2").setWrap(true);
  }
  // CRM
  const sh = ss.getSheetByName(CRM) || ss.insertSheet(CRM, 0);
  sh.getRange(1, 1, 1, CRM_COLS.length).setValues([CRM_COLS]).setFontWeight("bold")
    .setBackground("#252525").setFontColor("#ffffff");
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  [190, 220, 135, 120, 95, 110, 70, 120, 135, 135, 95, 120, 110, 280].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  const n = sh.getMaxRows() - 1;
  sh.getRange(2, 12, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(ETAPAS, true).setAllowInvalid(false).build());
  sh.getRange(2, 13, n, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireDate().setAllowInvalid(false).build()).setNumberFormat("dd/mm/yyyy");
  sh.getRange(2, 9, n, 2).setNumberFormat("dd/mm/yyyy hh:mm");
  const cores = { Novo: "#e3e3e3", Contatado: "#dcf6fd", Interessado: "#ece2ff", Negociando: "#fae1f0",
                  Comprou: "#c8f2d4", Perdido: "#d9d9d9" };
  const rules = Object.keys(cores).map((k) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(k).setBackground(cores[k]).setRanges([sh.getRange("L2:L")]).build());
  [["Quente", "#f72a9f", "#ffffff"], ["Morno", "#fae1f0", "#8a3e6b"], ["Frio", "#dcf6fd", "#0090b0"]].forEach((t) =>
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(t[0]).setBackground(t[1])
      .setFontColor(t[2]).setBold(true).setRanges([sh.getRange("K2:K")]).build()));
  rules.push(SpreadsheetApp.newConditionalFormatRule() // follow-up vencido: data em vermelho
    .whenFormulaSatisfied('=AND($M2<>"",$M2<=TODAY(),$L2<>"Comprou",$L2<>"Perdido")')
    .setFontColor("#c0003f").setBold(true).setRanges([sh.getRange("M2:M")]).build());
  sh.setConditionalFormatRules(rules);
  if (sh.getFilter() == null) sh.getRange(1, 1, sh.getMaxRows(), CRM_COLS.length).createFilter();
  reconstruirCRM(true);
  painel_();
  ss.setActiveSheet(sh);
}

// Lê todos os downloads da aba Leads e (re)monta as colunas automáticas do CRM.
// Etapa, próxima ação e observações já digitadas são preservadas.
function reconstruirCRM(silencioso) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CRM);
  if (!sh) return configurarCRM();
  const leads = aba_().getDataRange().getValues().slice(1);
  const manual = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues().forEach((r) => {
      if (r[1]) manual[String(r[1]).toLowerCase()] = [r[11] || "Novo", r[12], r[13]];
    });
    sh.getRange(2, 1, sh.getLastRow() - 1, 14).clearContent();
  }
  const pessoas = {}, ordem = [];
  leads.forEach((r) => {
    const email = String(r[2]).trim().toLowerCase();
    if (!email) return;
    let x = pessoas[email];
    if (!x) { x = pessoas[email] = { nome: r[1], zap: "", aceita: false, aulas: [], origem: r[5], ini: r[0], fim: r[0] }; ordem.push(email); }
    if (r[1]) x.nome = r[1];
    if (r[3]) x.zap = String(r[3]).replace(/[^\d+]/g, "");
    if (r[7] === "sim") x.aceita = true;
    const a = aula2_(r[4]);
    if (a && x.aulas.indexOf(a) < 0) x.aulas.push(a);
    if (r[0] > x.fim) x.fim = r[0];
  });
  const linhas = ordem.map((e) => {
    const x = pessoas[e], m = manual[e] || ["Novo", "", ""];
    x.aulas.sort();
    return [x.nome, e, x.zap ? "'" + x.zap : "", "", x.aceita ? "sim" : "não", x.aulas.join(", "),
            x.aulas.length, x.origem, x.ini, x.fim, "", m[0], m[1], m[2]];
  });
  if (linhas.length) {
    sh.getRange(2, 1, linhas.length, 14).setValues(linhas);
    for (let r = 2; r < 2 + linhas.length; r++) formulasLinha_(sh, r);
  }
  if (silencioso !== true) SpreadsheetApp.getActive().toast(linhas.length + " pessoas no CRM", "Método Lente");
}

function painel_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const p = ss.getSheetByName(PAINEL) || ss.insertSheet(PAINEL, 1);
  p.clear();
  p.getRange("A1").setValue("Método Lente · Painel de leads").setFontSize(18).setFontWeight("bold");
  const kpis = [
    ["Pessoas no CRM", "=COUNTA(CRM!B2:B)"],
    ["Downloads de mapas", "=COUNTA(Leads!C2:C)"],
    ["Aceitam WhatsApp", '=COUNTIF(CRM!E2:E,"sim")'],
    ["Leads quentes", '=COUNTIF(CRM!K2:K,"Quente")'],
    ["Follow-ups vencidos", '=COUNTIFS(CRM!M2:M,"<="&TODAY(),CRM!L2:L,"<>Comprou",CRM!L2:L,"<>Perdido")'],
    ["Compraram o curso", '=COUNTIF(CRM!L2:L,"Comprou")'],
    ["Conversão", "=IFERROR(B8/B3,0)"],
  ];
  p.getRange(3, 1, kpis.length, 2).setValues(kpis);
  p.getRange("B9").setNumberFormat("0.0%");
  p.getRange("A3:A9").setFontWeight("bold");
  p.getRange("B3:B9").setFontSize(14).setFontWeight("bold").setFontColor("#6400f8").setHorizontalAlignment("left");
  const blocos = [
    ["D2", "Funil (etapa)", '=QUERY(CRM!L2:L,"select L, count(L) where L is not null group by L label count(L) \'\'",0)'],
    ["G2", "Downloads por aula", '=QUERY(Leads!E2:E,"select E, count(E) where E is not null group by E label count(E) \'\'",0)'],
    ["J2", "Origem", '=QUERY(Leads!F2:F,"select F, count(F) where F is not null group by F label count(F) \'\'",0)'],
    ["D12", "Ligar / mandar mensagem hoje", '=IFERROR(SORT(FILTER({CRM!A2:A,CRM!L2:L,CRM!M2:M,CRM!D2:D},CRM!M2:M<>"",CRM!M2:M<=TODAY(),CRM!L2:L<>"Comprou",CRM!L2:L<>"Perdido"),3,TRUE),"Nada pendente")'],
    ["D24", "Quentes sem contato", '=IFERROR(FILTER({CRM!A2:A,CRM!F2:F,CRM!D2:D},CRM!K2:K="Quente",CRM!L2:L="Novo"),"Nenhum")'],
  ];
  blocos.forEach((b) => {
    const c = p.getRange(b[0]);
    c.setValue(b[1]).setFontWeight("bold").setBackground("#fae1f0");
    c.offset(1, 0).setFormula(b[2]);
  });
  p.setColumnWidth(1, 190); p.setColumnWidth(4, 190); p.setColumnWidth(7, 150); p.setColumnWidth(10, 120);
}
