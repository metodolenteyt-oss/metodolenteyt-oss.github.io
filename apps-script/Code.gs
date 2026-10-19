/**
 * Metodo Lente - captura de leads (gratis, roda no Google).
 * Cole este arquivo em Planilha > Extensoes > Apps Script e publique como "App da Web"
 * (Executar como: Eu - Quem pode acessar: Qualquer pessoa). Passo a passo no LEIA-ME.md.
 */
const SITE_URL = "https://metodolenteyt-oss.github.io/"; // troque apos publicar o site
const ABA = "Leads";
const COLUNAS = ["data", "nome", "email", "whatsapp", "aula", "origem", "consent_email",
                 "consent_whatsapp", "idade_ok", "retorno", "pagina", "ua"];

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.site) return ok_(); // honeypot preenchido = robo
  const email = String(p.email || "").trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return ok_();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = aba_();
    const agora = new Date();
    sh.appendRow(COLUNAS.map((c) => (c === "data" ? agora : c === "email" ? email
      : c === "aula" ? "'" + aula2_(p.aula) : limpa_(p[c]))));
    try { crmUpsert_(email, p, agora); } catch (err) { console.error("CRM", err); } // o lead ja esta salvo
  } finally {
    lock.releaseLock();
  }

  if (p.retorno !== "sim") enviaEmail_(email, limpa_(p.nome), limpa_(p.aula));
  return ok_();
}

function doGet() {
  return ContentService.createTextOutput("M\u00E9todo Lente \u2014 captura ativa \u2714");
}

function enviaEmail_(email, nome, aula) {
  const link = SITE_URL + "?a=" + encodeURIComponent(aula) + "&baixar=1";
  const prev = SITE_URL + "previews/aula-" + encodeURIComponent(aula) + ".jpg";
  const logo = SITE_URL + "assets/logo-colorido.png";
  let primeiro = (nome || "").split(" ")[0] || "";
  primeiro = primeiro.charAt(0).toUpperCase() + primeiro.slice(1).toLowerCase();
  const serif = "Georgia,'Times New Roman',serif", sans = "Arial,Helvetica,sans-serif";
  const html =
    '<div style="background:#e3e3e3;padding:28px 12px;font-family:' + sans + ';color:#252525">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;margin:0 auto;background:#ffffff;border-radius:14px;border-top:8px solid #f72a9f">' +
    '<tr><td style="padding:26px 30px 6px;text-align:center">' +
    '<img src="' + logo + '" width="190" alt="Lente - Profe Lu" style="display:inline-block;border:0;width:190px;height:auto"></td></tr>' +
    '<tr><td style="padding:18px 30px 4px;font-family:' + serif + ';font-size:28px;line-height:1.15;font-weight:bold">' +
    'Oi, ' + primeiro + '! Seu mapa mental chegou.</td></tr>' +
    '<tr><td style="padding:8px 30px 18px;font-size:15px;line-height:1.5">' +
    'Obrigada por assistir &agrave; aula. Aqui est&aacute; o mapa mental para revisar tudo o que vimos, num s&oacute; lugar.</td></tr>' +
    '<tr><td style="padding:0 30px;text-align:center"><a href="' + link + '">' +
    '<img src="' + prev + '" width="500" alt="Pr&eacute;via do mapa mental" style="width:100%;max-width:500px;height:auto;border-radius:6px;border:1px solid #e3e3e3"></a></td></tr>' +
    '<tr><td style="padding:22px 30px 8px;text-align:center">' +
    '<a href="' + link + '" style="display:inline-block;background:#f72a9f;color:#ffffff;padding:15px 30px;border-radius:30px;text-decoration:none;font-weight:bold;font-size:16px">' +
    'Baixar o mapa mental (PDF)</a></td></tr>' +
    '<tr><td style="padding:10px 30px 26px;text-align:center;font-size:14px;color:#8a3e6b;font-family:' + serif + ';font-style:italic">' +
    'Nos pr&oacute;ximos v&iacute;deos tem mais mapa mental. Fique de olho no canal!</td></tr>' +
    '</table>' +
    '<p style="max-width:560px;margin:14px auto 0;font-size:11px;line-height:1.5;color:#666;text-align:center">' +
    'Voc&ecirc; recebeu este e-mail porque pediu o mapa mental no site do M&eacute;todo Lente. ' +
    'Para n&atilde;o receber mais, responda com a palavra SAIR.</p></div>';
  try {
    MailApp.sendEmail({
      to: email,
      name: "Profe Lu | Lente",
      subject: "Seu mapa mental da aula " + aula + " chegou",
      htmlBody: html,
      body: "Oi, " + primeiro + "! Baixe seu mapa mental: " + link,
    });
  } catch (err) {
    console.error("e-mail nao enviado (cota diaria?)", err); // o lead ja foi salvo
  }
}

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(ABA) || ss.insertSheet(ABA);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUNAS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, COLUNAS.length).setFontWeight("bold").setBackground("#fae1f0");
  }
  return sh;
}

function limpa_(v) {
  // evita injecao de formula na planilha (=, +, -, @)
  const s = String(v == null ? "" : v).slice(0, 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

// ===================================================================== CRM
// Aba CRM: 1 linha por pessoa (chave = e-mail). Colunas AJ sao preenchidas pelo script;
// LN (Etapa, Proxima acao, Observacoes) sao da equipe e NUNCA sao sobrescritas.
const CRM = "CRM", PAINEL = "Painel", CONFIG = "Config";
const CRM_COLS = ["Nome", "E-mail", "WhatsApp", "Abrir WhatsApp", "Aceita WhatsApp", "Aulas baixadas",
                  "N\u00BA de aulas", "Origem (1\u00BA contato)", "Primeiro contato", "\u00DAltima atividade",
                  "Temperatura", "Etapa", "Pr\u00F3xima a\u00E7\u00E3o", "Observa\u00E7\u00F5es"];
const ETAPAS = ["Novo", "Contatado", "Interessado", "Negociando", "Comprou", "Perdido"];
const MSG_PADRAO = "Oi, {nome}! Aqui \u00E9 a Profe Lu, do M\u00E9todo Lente. Vi que voc\u00EA baixou o mapa mental " +
                   "da aula {aula}. Posso te contar como funciona o curso completo?";

function onOpen() {
  SpreadsheetApp.getUi().createMenu("M\u00E9todo Lente")
    .addItem("Configurar / atualizar CRM", "configurarCRM")
    .addItem("Reconstruir CRM a partir de Leads", "reconstruirCRM")
    .addItem("Atualizar painel", "painel_")
    .addToUi();
}

function aula2_(v) {
  const s = String(v == null ? "" : v).replace(/\D/g, "");
  return s ? ("0" + s).slice(-2) : "";
}

function crmUpsert_(email, p, quando) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CRM);
  if (!sh) return; // CRM ainda nao configurado
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
    sh.appendRow([limpa_(p.nome), email, zap ? "'" + zap : "", "", aceita ? "sim" : "n\u00E3o", aula,
                  aula ? 1 : 0, limpa_(p.origem), quando, quando, "", "Novo", "", ""]);
    formulasLinha_(sh, sh.getLastRow());
  }
}

function formulasLinha_(sh, r) {
  // D: link do WhatsApp com a mensagem da aba Config (so para quem autorizou)
  sh.getRange(r, 4).setFormulaR1C1(
    '=IF(RC2="","",IF(RC5<>"sim","sem autoriza\u00E7\u00E3o",HYPERLINK("https://wa.me/"&REGEXREPLACE(RC3,"\\D","")&"?text="&' +
    'ENCODEURL(SUBSTITUTE(SUBSTITUTE(Config!R2C2,"{nome}",PROPER(REGEXEXTRACT(RC1&" ","^\\S+"))),"{aula}",' +
    'REGEXEXTRACT(RC6&"","[^,]*$"))),"Abrir conversa")))');
  // K: temperatura = no de aulas + quao recente foi a ultima atividade
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
    cf.getRange("A1:B1").setValues([["Configura\u00E7\u00E3o", "Valor"]]).setFontWeight("bold").setBackground("#fae1f0");
    cf.getRange("A2:B2").setValues([["Mensagem de WhatsApp ({nome} e {aula} s\u00E3o trocados automaticamente)", MSG_PADRAO]]);
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
  sh.getBandings().forEach((b) => b.remove());
  sh.getRange(1, 1, sh.getMaxRows(), CRM_COLS.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor("#252525").setFirstRowColor("#ffffff").setSecondRowColor("#f7f5fa");
  sh.getRange("A:N").setFontFamily("Quicksand").setVerticalAlignment("middle");
  sh.setRowHeight(1, 34);
  reconstruirCRM(true);
  painel_();
  ss.setActiveSheet(sh);
}

// Le todos os downloads da aba Leads e (re)monta as colunas automaticas do CRM.
// Etapa, proxima acao e observacoes ja digitadas sao preservadas.
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
    return [x.nome, e, x.zap ? "'" + x.zap : "", "", x.aceita ? "sim" : "n\u00E3o", x.aulas.join(", "),
            x.aulas.length, x.origem, x.ini, x.fim, "", m[0], m[1], m[2]];
  });
  if (linhas.length) {
    sh.getRange(2, 1, linhas.length, 14).setValues(linhas);
    for (let r = 2; r < 2 + linhas.length; r++) formulasLinha_(sh, r);
  }
  if (silencioso !== true) SpreadsheetApp.getActive().toast(linhas.length + " pessoas no CRM", "M\u00E9todo Lente");
}

function painel_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const p = ss.getSheetByName(PAINEL) || ss.insertSheet(PAINEL, 1);
  p.getCharts().forEach((c) => p.removeChart(c));
  p.clear(); p.clearFormats();
  p.getBandings().forEach((b) => b.remove());
  if (p.getMaxColumns() < 24) p.insertColumnsAfter(p.getMaxColumns(), 24 - p.getMaxColumns());
  p.setHiddenGridlines(true);
  const F = "Quicksand", INK = "#252525", BG = "#f4f4f6";
  p.getRange("A1:L80").setBackground(BG).setFontFamily(F).setFontColor(INK).setVerticalAlignment("middle");
  for (let c = 1; c <= 12; c++) p.setColumnWidth(c, 112);
  p.setColumnWidth(13, 24);

  // faixa de titulo
  p.setRowHeight(1, 46); p.setRowHeight(2, 24);
  p.getRange("A1:L2").setBackground(INK);
  p.getRange("A1:H1").merge().setValue("M\u00C9TODO LENTE  \u00B7  PAINEL DE LEADS")
    .setFontSize(18).setFontWeight("bold").setFontColor("#ffffff");
  p.getRange("A2:H2").merge().setValue("Capta\u00E7\u00E3o pelos mapas mentais das aulas e funil de venda do curso")
    .setFontSize(10).setFontColor("#cfcfcf");
  p.getRange("I1:L2").merge().setFormula('="Atualizado em "&TEXT(NOW(),"dd/mm/yyyy hh:mm")')
    .setFontSize(10).setFontColor("#cfcfcf").setHorizontalAlignment("right");
  p.getRange("A1:L2").setBorder(null, null, true, null, null, null, "#f72a9f", SpreadsheetApp.BorderStyle.SOLID_THICK);

  // cartoes de KPI (2 colunas cada)
  const kpis = [
    ["PESSOAS NO CRM", "=COUNTA(CRM!B2:B)", "0", "#6400f8"],
    ["DOWNLOADS", "=COUNTA(Leads!C2:C)", "0", "#01c8f1"],
    ["ACEITAM WHATSAPP", '=COUNTIF(CRM!E2:E,"sim")', "0", "#25d366"],
    ["LEADS QUENTES", '=COUNTIF(CRM!K2:K,"Quente")', "0", "#f72a9f"],
    ["FOLLOW-UPS VENCIDOS", '=COUNTIFS(CRM!M2:M,"<="&TODAY(),CRM!L2:L,"<>Comprou",CRM!L2:L,"<>Perdido")', "0", "#c0003f"],
    ["CONVERS\u00C3O EM VENDA", '=IFERROR(COUNTIF(CRM!L2:L,"Comprou")/COUNTA(CRM!B2:B),0)', "0.0%", "#8a3e6b"],
  ];
  p.setRowHeight(3, 14); p.setRowHeight(4, 26); p.setRowHeight(5, 30); p.setRowHeight(6, 30); p.setRowHeight(7, 22);
  kpis.forEach((k, i) => {
    const c = 1 + i * 2;
    const card = p.getRange(4, c, 4, 2);
    card.setBackground("#ffffff")
      .setBorder(true, true, true, true, null, null, "#e1e1e6", SpreadsheetApp.BorderStyle.SOLID);
    p.getRange(4, c, 1, 2).setBorder(true, null, null, null, null, null, k[3], SpreadsheetApp.BorderStyle.SOLID_THICK);
    p.getRange(4, c, 1, 2).merge().setValue(k[0]).setFontSize(9).setFontWeight("bold").setFontColor("#7a7a85")
      .setHorizontalAlignment("center");
    p.getRange(5, c, 2, 2).merge().setFormula(k[1]).setNumberFormat(k[2]).setFontSize(28).setFontWeight("bold")
      .setFontColor(k[3]).setHorizontalAlignment("center");
  });
  p.getRange("A7:B7").merge().setFormula('="+"&COUNTIF(CRM!I2:I,">="&TODAY()-7)&" nos \u00FAltimos 7 dias"');
  p.getRange("C7:D7").merge().setFormula('=IFERROR(TEXT(COUNTA(Leads!C2:C)/COUNTA(CRM!B2:B),"0.0")&" mapas por pessoa","")');
  p.getRange("E7:F7").merge().setFormula('=IFERROR(TEXT(COUNTIF(CRM!E2:E,"sim")/COUNTA(CRM!B2:B),"0%")&" do total","")');
  p.getRange("G7:H7").merge().setFormula('=COUNTIFS(CRM!K2:K,"Quente",CRM!L2:L,"Novo")&" ainda sem contato"');
  p.getRange("I7:J7").merge().setValue("agenda vencida");
  p.getRange("K7:L7").merge().setFormula('=COUNTIF(CRM!L2:L,"Comprou")&" alunos"');
  p.getRange("A7:L7").setFontSize(9).setFontColor("#7a7a85").setHorizontalAlignment("center");

  // dados auxiliares dos graficos (colunas N:W, ocultas)
  p.getRange("N1").setFormula('=QUERY(Leads!A2:C,"select toDate(A), count(C) where C is not null group by toDate(A) label toDate(A) \'Dia\', count(C) \'Leads\'",0)');
  p.getRange("P1:Q1").setValues([["Etapa", "Pessoas"]]);
  ETAPAS.forEach((e, i) => p.getRange(2 + i, 16, 1, 2).setValues([[e, '=COUNTIF(CRM!L2:L,"' + e + '")']]));
  p.getRange("S1").setFormula('=QUERY(Leads!E2:E,"select E, count(E) where E is not null group by E label E \'Aula\', count(E) \'Downloads\'",0)');
  p.getRange("U1").setFormula('=QUERY(Leads!F2:F,"select F, count(F) where F is not null group by F label F \'Origem\', count(F) \'Downloads\'",0)');
  p.getRange("W1:X1").setValues([["Temperatura", "Pessoas"]]);
  ["Quente", "Morno", "Frio"].forEach((t, i) => p.getRange(2 + i, 23, 1, 2).setValues([[t, '=COUNTIF(CRM!K2:K,"' + t + '")']]));
  p.hideColumns(14, 11);

  // graficos
  const txt = { color: INK, fontName: F };
  const base = (b, title, w, h) => b.setOption("title", title).setOption("titleTextStyle", { color: INK, fontName: F, fontSize: 13, bold: true })
    .setOption("backgroundColor", { fill: "#ffffff", stroke: "#e1e1e6", strokeWidth: 1 })
    .setOption("legend", { position: "none" }).setOption("width", w).setOption("height", h)
    .setOption("hAxis", { textStyle: txt, gridlines: { color: "#eeeeee" } })
    .setOption("vAxis", { textStyle: txt, gridlines: { color: "#eeeeee" }, minValue: 0 });
  p.setRowHeight(8, 14);
  p.insertChart(base(p.newChart().setChartType(Charts.ChartType.AREA).addRange(p.getRange("N1:O400")), "Leads por dia", 806, 300)
    .setOption("colors", ["#6400f8"]).setOption("areaOpacity", 0.15).setOption("lineWidth", 3).setOption("pointSize", 5)
    .setPosition(9, 1, 0, 0).build());
  p.insertChart(base(p.newChart().setChartType(Charts.ChartType.PIE).addRange(p.getRange("W1:X4")), "Temperatura", 520, 300)
    .setOption("pieHole", 0.55).setOption("colors", ["#f72a9f", "#fae1f0", "#01c8f1"])
    .setOption("legend", { position: "right", textStyle: txt }).setOption("pieSliceTextStyle", { color: INK, fontName: F })
    .setPosition(9, 8, 0, 0).build());
  p.insertChart(base(p.newChart().setChartType(Charts.ChartType.BAR).addRange(p.getRange("P1:Q7")), "Funil de vendas (pessoas por etapa)", 806, 300)
    .setOption("colors", ["#f72a9f"]).setOption("dataLabel", "value").setOption("bar", { groupWidth: "70%" })
    .setPosition(25, 1, 0, 0).build());
  p.insertChart(base(p.newChart().setChartType(Charts.ChartType.PIE).addRange(p.getRange("U1:V20")), "Origem dos downloads", 520, 300)
    .setOption("pieHole", 0.55).setOption("colors", ["#6400f8", "#01c8f1", "#f72a9f", "#8a3e6b"])
    .setOption("legend", { position: "right", textStyle: txt }).setPosition(25, 8, 0, 0).build());
  p.insertChart(base(p.newChart().setChartType(Charts.ChartType.COLUMN).addRange(p.getRange("S1:T60")), "Downloads por aula", 1340, 280)
    .setOption("colors", ["#01c8f1"]).setOption("dataLabel", "value").setPosition(41, 1, 0, 0).build());

  // listas de acao
  const lista = (row, titulo, cab, formula, cor) => {
    p.getRange(row, 1, 1, 12).merge().setValue(titulo).setFontSize(13).setFontWeight("bold")
      .setBackground("#ffffff").setBorder(true, null, null, null, null, null, cor, SpreadsheetApp.BorderStyle.SOLID_THICK);
    p.getRange(row + 1, 1, 1, cab.length).setValues([cab]).setFontWeight("bold").setFontSize(9)
      .setBackground(INK).setFontColor("#ffffff");
    p.getRange(row + 2, 1).setFormula(formula);
    p.getRange(row + 2, 1, 12, 12).setBackground("#ffffff").setFontSize(10);
  };
  lista(57, "Ligar / mandar mensagem hoje", ["Nome", "Etapa", "Pr\u00F3xima a\u00E7\u00E3o", "WhatsApp"],
    '=IFERROR(SORT(FILTER({CRM!A2:A,CRM!L2:L,CRM!M2:M,CRM!D2:D},CRM!M2:M<>"",CRM!M2:M<=TODAY(),CRM!L2:L<>"Comprou",CRM!L2:L<>"Perdido"),3,TRUE),"Nada pendente")', "#c0003f");
  p.getRange("C59:C70").setNumberFormat("dd/mm/yyyy");
  lista(72, "Leads quentes ainda sem contato", ["Nome", "Aulas", "\u00DAltima atividade", "WhatsApp"],
    '=IFERROR(SORT(FILTER({CRM!A2:A,CRM!F2:F,CRM!J2:J,CRM!D2:D},CRM!K2:K="Quente",CRM!L2:L="Novo"),3,FALSE),"Nenhum")', "#f72a9f");
  p.getRange("C74:C85").setNumberFormat("dd/mm/yyyy");
  p.setFrozenRows(2);
}
