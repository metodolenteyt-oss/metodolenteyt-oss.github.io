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
    sh.appendRow(COLUNAS.map((c) => (c === "data" ? new Date() : c === "email" ? email : limpa_(p[c]))));
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
