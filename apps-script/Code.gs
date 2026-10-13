/**
 * Método Lente — captura de leads (grátis, roda no Google).
 * Cole este arquivo em Planilha > Extensões > Apps Script e publique como "App da Web"
 * (Executar como: Eu · Quem pode acessar: Qualquer pessoa). Passo a passo no LEIA-ME.md.
 */
const SITE_URL = "https://metodolenteyt-oss.github.io/"; // troque após publicar o site
const ABA = "Leads";
const COLUNAS = ["data", "nome", "email", "whatsapp", "aula", "origem", "consent_email",
                 "consent_whatsapp", "idade_ok", "retorno", "pagina", "ua"];

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.site) return ok_(); // honeypot preenchido = robô
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
  return ContentService.createTextOutput("Método Lente — captura ativa ✔");
}

function enviaEmail_(email, nome, aula) {
  const link = SITE_URL + "?a=" + encodeURIComponent(aula) + "&baixar=1";
  const primeiro = (nome || "").split(" ")[0] || "";
  try {
    MailApp.sendEmail({
      to: email,
      name: "Profe Lu · Lente",
      subject: "Seu mapa mental da aula " + aula + " chegou 📚",
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:520px;margin:auto;color:#252525">' +
        '<p style="color:#6400f8;letter-spacing:3px;font-size:12px;font-weight:bold">LENTE · LEITURA DOS CLÁSSICOS ALÉM DO ÓBVIO</p>' +
        "<h2>Oi, " + primeiro + "! Seu mapa mental está aqui.</h2>" +
        "<p>Obrigada por assistir à aula. Clique para baixar o PDF:</p>" +
        '<p><a href="' + link + '" style="background:#f72a9f;color:#fff;padding:14px 22px;border-radius:30px;text-decoration:none;font-weight:bold">Baixar o mapa mental</a></p>' +
        "<p>Nos próximos vídeos tem mais mapa — fique de olho no canal.</p>" +
        "<p>Beijo,<br><b>Profe Lu</b></p>" +
        '<p style="font-size:11px;color:#777">Você recebeu este e-mail porque pediu o mapa mental no site do Método Lente. ' +
        "Para não receber mais, responda com a palavra SAIR.</p></div>",
    });
  } catch (err) {
    console.error("e-mail não enviado (cota diária?)", err); // o lead já foi salvo
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
  // evita injeção de fórmula na planilha (=, +, -, @)
  const s = String(v == null ? "" : v).slice(0, 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function ok_() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}
