(() => {
  const cfg = window.LENTE;
  const q = new URLSearchParams(location.search);
  const ids = Object.keys(cfg.AULAS);
  const id = cfg.AULAS[q.get("a")] ? q.get("a") : ids[ids.length - 1];
  const aula = cfg.AULAS[id];
  const $ = (s) => document.querySelector(s);

  $("#kicker").textContent = aula.kicker;
  $("#titulo").textContent = aula.titulo;
  $("#sub").textContent = aula.sub;
  $("#preview").src = aula.preview;
  $("#dl").href = aula.pdf;
  $("#canal").href = cfg.CANAL_URL;
  document.title = aula.titulo + " · Lente — Profe Lu";

  const done = (nome) => {
    $("#lead-form").hidden = true;
    $("#ok-nome").textContent = (nome || "").split(" ")[0] || "tudo certo";
    $("#ok").hidden = false;
  };
  // link do e-mail (…?a=01&baixar=1) ou quem já se cadastrou neste aparelho vai direto ao download
  let saved = null;
  try { saved = JSON.parse(localStorage.getItem("lente-lead") || "null"); } catch (e) {}
  const send = (fields) => {
    const body = new URLSearchParams({ aula: id, origem: q.get("o") || "direto", pagina: location.href.split("?")[0],
      ua: navigator.userAgent.slice(0, 120), ...fields });
    if (!cfg.APPS_SCRIPT_URL) return console.warn("APPS_SCRIPT_URL vazio em aulas.js — o lead NÃO foi salvo (modo teste).");
    return fetch(cfg.APPS_SCRIPT_URL, { method: "POST", mode: "no-cors", body }).catch(console.error);
  };
  if (q.get("baixar") === "1") done(saved && saved.nome);
  else if (saved && saved.email) {
    // já cadastrado neste aparelho: entrega direto e registra o download desta aula (1x por aula)
    done(saved.nome);
    if (!(saved.aulas || []).includes(id)) {
      send({ ...saved.dados, retorno: "sim" });
      saved.aulas = [...(saved.aulas || []), id];
      try { localStorage.setItem("lente-lead", JSON.stringify(saved)); } catch (e) {}
    }
  }

  // máscara de WhatsApp: (48) 99999-9999
  const wa = $("input[name=whatsapp]");
  wa.addEventListener("input", () => {
    const d = wa.value.replace(/\D/g, "").slice(0, 11);
    wa.value = d.length > 6 ? `(${d.slice(0, 2)}) ${d.slice(2, d.length - 4)}-${d.slice(-4)}`
      : d.length > 2 ? `(${d.slice(0, 2)}) ${d.slice(2)}` : d;
  });

  $("#lead-form").addEventListener("submit", async (ev) => {
    ev.preventDefault();
    const f = ev.target, err = $("#err");
    const nome = f.nome.value.trim().replace(/\s+/g, " ");
    const email = f.email.value.trim().toLowerCase();
    const fone = f.whatsapp.value.replace(/\D/g, "");
    err.textContent = "";
    if (nome.split(" ").length < 2) return (err.textContent = "Digite seu nome completo.");
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return (err.textContent = "Confira o e-mail.");
    if (fone.length < 10) return (err.textContent = "Digite o WhatsApp com DDD.");
    if (!f.consent.checked || !f.idade.checked) return (err.textContent = "Marque as caixas obrigatórias para continuar.");

    $("#btn").disabled = true;
    $("#btn").textContent = "Enviando…";
    const dados = { nome, email, whatsapp: "+55" + fone, consent_email: "sim",
      consent_whatsapp: f.whats_ok.checked ? "sim" : "não", idade_ok: "sim" };
    await send({ ...dados, site: f.site.value }); // mesmo se falhar, a pessoa recebe o mapa
    try { localStorage.setItem("lente-lead", JSON.stringify({ nome, email, dados, aulas: [id] })); } catch (e) {}
    done(nome);
  });
})();
