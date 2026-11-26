/**
 * KP CHECKLIST — servidor (Google Apps Script)
 * Recebe os check-lists do app, grava na planilha e guarda as fotos no Drive.
 *
 * ANTES DE USAR: troque as duas senhas abaixo, salve e rode a função "configurar" uma vez.
 */
const SENHA_ADMINISTRATIVO = "TROQUE-esta-senha-ADM";
const SENHA_DIRETORIA = "TROQUE-esta-senha-DIR";
const CODIGO_EMPRESA = "KP-7Q4M-2026"; // igual ao CODIGO do arquivo config.js do app

const ABA = "Checklists";
const FUSO = "America/Sao_Paulo";
const COLS = ["ID", "Situação", "Data saída", "Hora saída", "Placa", "Motorista", "CPF", "Ajudante", "Ajudante 2",
  "Km saída", "Nº transporte", "Lacre", "Qtd. clientes", "Peso carga", "Itens com problema (NOK)",
  "Baú refrigerado", "Temp. exigida", "Temp. saída", "Aparelho frio saída", "Avarias saída", "Obs. saída",
  "Data retorno", "Hora retorno", "Km chegada", "Km rodado", "Clientes retornados", "Peso retornado",
  "Temp. retorno", "Aparelho frio retorno", "Avarias retorno", "Obs. retorno", "Pasta de fotos",
  "Validado por", "Validado em", "Comentário validação", "Recebido em", "dados_saida", "dados_retorno"];
const c = nome => COLS.indexOf(nome) + 1;

/* ---------- configuração (rodar uma vez) ---------- */
function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Abra este script pela planilha: Extensões > Apps Script.");
  const P = PropertiesService.getScriptProperties();
  let sh = ss.getSheetByName(ABA);
  if (!sh) sh = ss.insertSheet(ABA, 0);
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight("bold").setBackground("#0a5cff")
    .setFontColor("#ffffff").setWrap(true).setVerticalAlignment("middle");
  sh.setFrozenRows(1);
  sh.setRowHeight(1, 42);
  if (sh.getMaxRows() < 15000) sh.insertRowsAfter(sh.getMaxRows(), 15000 - sh.getMaxRows());
  const n = sh.getMaxRows() - 1;
  ["Data saída", "Data retorno"].forEach(k => sh.getRange(2, c(k), n).setNumberFormat("dd/mm/yyyy"));
  ["Hora saída", "Hora retorno", "CPF", "Placa", "ID"].forEach(k => sh.getRange(2, c(k), n).setNumberFormat("@"));
  ["Validado em", "Recebido em"].forEach(k => sh.getRange(2, c(k), n).setNumberFormat("dd/mm/yyyy hh:mm"));
  sh.hideColumns(c("dados_saida"), 2);
  sh.setColumnWidth(c("ID"), 60);
  sh.setColumnWidths(c("Itens com problema (NOK)"), 1, 260);
  // cores automáticas
  const tudo = r => sh.getRange(2, c(r), n);
  const R = SpreadsheetApp.newConditionalFormatRule;
  sh.setConditionalFormatRules([
    R().whenTextEqualTo("Validado").setBackground("#c8eed9").setFontColor("#05663a").setRanges([tudo("Situação")]).build(),
    R().whenTextEqualTo("Reprovado").setBackground("#e6261f").setFontColor("#ffffff").setRanges([tudo("Situação")]).build(),
    R().whenTextEqualTo("Em rota").setBackground("#ffe7a3").setRanges([tudo("Situação")]).build(),
    R().whenTextEqualTo("Aguardando validação").setBackground("#d6e4ff").setRanges([tudo("Situação")]).build(),
    R().whenTextEqualTo("Retorno sem saída").setBackground("#eeeeee").setRanges([tudo("Situação")]).build(),
    R().whenCellNotEmpty().setBackground("#fdecea").setFontColor("#c0150e").setBold(true)
      .setRanges([tudo("Itens com problema (NOK)"), tudo("Avarias saída"), tudo("Avarias retorno")]).build(),
    R().whenTextEqualTo("COM PROBLEMA").setBackground("#e6261f").setFontColor("#ffffff")
      .setRanges([tudo("Aparelho frio saída"), tudo("Aparelho frio retorno")]).build(),
    R().whenTextContains("FORA").setBackground("#fdecea").setFontColor("#c0150e").setBold(true)
      .setRanges([tudo("Temp. saída"), tudo("Temp. retorno")]).build()
  ]);
  ["Página1", "Sheet1"].forEach(nm => { const x = ss.getSheetByName(nm); if (x && x.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(x); });
  if (!P.getProperty("PASTA_ID")) {
    const pasta = DriveApp.createFolder("KP CHECKLIST - Fotos");
    P.setProperty("PASTA_ID", pasta.getId());
  }
  if (!P.getProperty("SEGREDO")) P.setProperty("SEGREDO", Utilities.getUuid() + Utilities.getUuid());
  P.setProperty("PLANILHA_ID", ss.getId());
  Logger.log("Pronto! Agora publique: Implantar > Nova implantação > App da Web.");
}

/* ---------- entrada ---------- */
function doGet() { return json({ ok: true, app: "KP CHECKLIST" }); }
function doPost(e) {
  let q;
  try { q = JSON.parse(e.postData.contents); } catch (x) { return json({ ok: false, erro: "formato" }); }
  try {
    switch (q.acao) {
      case "enviar":
        if (q.codigo !== CODIGO_EMPRESA) return json({ ok: false, erro: "codigo" });
        return json(receber(q.registro));
      case "login": return json(login(q));
      case "listar": return json(listar(q));
      case "foto": return json(foto(q));
      case "validar": return json(validar(q));
    }
    return json({ ok: false, erro: "acao" });
  } catch (err) {
    const m = String(err && err.message || err);
    return json({ ok: false, erro: m === "login" ? "login" : m });
  }
}
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- utilidades ---------- */
function props() { return PropertiesService.getScriptProperties(); }
function aba() {
  const id = props().getProperty("PLANILHA_ID");
  if (!id) throw new Error("Rode a função configurar primeiro.");
  return SpreadsheetApp.openById(id).getSheetByName(ABA);
}
function txt(v, max) {
  let s = v == null ? "" : String(v).slice(0, max || 500);
  if (/^[=+\-@]/.test(s) && isNaN(Number(s))) s = "'" + s; // evita fórmulas digitadas
  return s;
}
function numero(v) { const n = Number(v); return v === null || v === "" || isNaN(n) ? "" : n; }
function placaOk(p) { return String(p || "").toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 8); }
function dataDe(s) { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || "")); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : ""; }
function iso(d) { return d instanceof Date ? Utilities.formatDate(d, FUSO, "yyyy-MM-dd") : String(d || ""); }
function horaTxt(v) { return v instanceof Date ? Utilities.formatDate(v, FUSO, "HH:mm") : String(v || ""); }
function fixarTexto(sh, n, L) {
  ["ID", "Hora saída", "CPF", "Placa"].forEach(k => { const x = sh.getRange(n, c(k)); x.setNumberFormat("@"); x.setValue(L[c(k) - 1]); });
}
function linhaDoId(sh, id) {
  if (!id || sh.getLastRow() < 2) return 0;
  const f = sh.getRange(2, c("ID"), sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}
function linhaAbertaDaPlaca(sh, placa) {
  const n = sh.getLastRow() - 1; if (n < 1) return 0;
  const pl = sh.getRange(2, c("Placa"), n, 1).getValues(), st = sh.getRange(2, c("Situação"), n, 1).getValues();
  for (let i = n - 1; i >= 0; i--) if (String(pl[i][0]) === placa && st[i][0] === "Em rota") return i + 2;
  return 0;
}
function faixaTxt(f) { return f && (f.exigMin != null || f.exigMax != null) ? (f.exigMin != null ? f.exigMin : "") + " a " + (f.exigMax != null ? f.exigMax : "") + " °C" : ""; }
function fora(t, f) { return t != null && f && ((f.exigMin != null && t < f.exigMin) || (f.exigMax != null && t > f.exigMax)); }
function tempTxt(f) { if (!f || f.temp == null) return ""; return f.temp + (fora(f.temp, f) ? " (FORA DA FAIXA)" : ""); }
function limpo(o, max) { const s = String(o == null ? "" : o); return s.length > (max || 200) ? s.slice(0, max || 200) : s; }

/* ---------- fotos ---------- */
function pastaDoChecklist(r, existenteId) {
  if (existenteId) { try { return DriveApp.getFolderById(existenteId); } catch (e) { } }
  const raiz = DriveApp.getFolderById(props().getProperty("PASTA_ID"));
  const mes = String(r.data || "").slice(0, 7) || "sem-data";
  const it = raiz.getFoldersByName(mes);
  const pm = it.hasNext() ? it.next() : raiz.createFolder(mes);
  return pm.createFolder((r.data || "") + " " + placaOk(r.placa) + " " + limpo(r.motorista && r.motorista.nome, 40));
}
function salvarImagem(pasta, b64, nome, tipo) {
  if (!b64 || typeof b64 !== "string" || b64.length > 15000000) return "";
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), tipo || "image/jpeg", "KP_" + nome);
  return pasta.createFile(blob).getId();
}
function salvarFotos(pasta, r, fase) {
  const pl = placaOk(r.placa), ids = {}, av = [];
  Object.keys(r.fotos || {}).slice(0, 10).forEach(k => {
    ids[k] = salvarImagem(pasta, r.fotos[k], pl + "_" + fase + "_" + String(k).replace(/[^a-z]/g, "") + ".jpg");
  });
  (r.avarias || []).slice(0, 20).forEach((a, i) => {
    av.push({ desc: limpo(a.desc, 200), foto: salvarImagem(pasta, a.foto, pl + "_" + fase + "_avaria" + (i + 1) + ".jpg") });
  });
  const ass = r.assinatura ? salvarImagem(pasta, r.assinatura, pl + "_assinatura.png", "image/png") : "";
  return { ids, av, ass };
}

/* ---------- receber check-list ---------- */
function receber(r) {
  if (!r || !r.id || (r.tipo !== "saida" && r.tipo !== "retorno")) return { ok: false, erro: "registro" };
  r.placa = placaOk(r.placa);
  const sh = aba();
  if (r.tipo === "saida") {
    if (linhaDoId(sh, r.id)) return { ok: true, repetido: true };
    const pasta = pastaDoChecklist(r);
    const f = salvarFotos(pasta, r, "saida");
    const lock = LockService.getScriptLock(); lock.waitLock(30000);
    try {
      if (linhaDoId(sh, r.id)) return { ok: true, repetido: true };
      const fr = r.frio, nok = r.itensNok || {};
      const dados = { km: r.km, transporte: r.transporte, lacre: r.lacre, clientes: r.clientes, peso: r.peso, itensNok: nok, obs: r.obs, gps: r.gps,
        frio: fr ? { temp: fr.temp, exigMin: fr.exigMin, exigMax: fr.exigMax, aparelho: fr.aparelho, fora: fora(fr.temp, fr) } : null,
        fotosIds: f.ids, avariasIds: f.av.map(a => ({ desc: a.desc, foto: a.foto })), pastaId: pasta.getId(), criadoEm: r.criadoEm, versao: r.versao };
      const L = new Array(COLS.length).fill("");
      const put = (k, v) => { L[c(k) - 1] = v; };
      put("ID", r.id); put("Situação", "Em rota"); put("Data saída", dataDe(r.data)); put("Hora saída", txt(r.hora, 5));
      put("Placa", r.placa); put("Motorista", txt(r.motorista && r.motorista.nome, 80)); put("CPF", txt(r.motorista && r.motorista.cpf, 14));
      put("Ajudante", txt(r.ajudante, 80)); put("Ajudante 2", txt(r.ajudante2, 80)); put("Km saída", numero(r.km));
      put("Nº transporte", txt(r.transporte, 40)); put("Lacre", txt(r.lacre, 40)); put("Qtd. clientes", numero(r.clientes)); put("Peso carga", txt(r.peso, 20));
      put("Itens com problema (NOK)", txt(Object.keys(nok).map(k => k + ": " + nok[k]).join("; "), 2000));
      put("Baú refrigerado", fr ? "Sim" : "Não"); put("Temp. exigida", txt(faixaTxt(fr))); put("Temp. saída", txt(tempTxt(fr)));
      put("Aparelho frio saída", fr ? (fr.aparelho === "nok" ? "COM PROBLEMA" : "OK") : "");
      put("Avarias saída", txt(f.av.map((a, i) => (i + 1) + ". " + a.desc).join("; "), 2000)); put("Obs. saída", txt(r.obs, 2000));
      put("Pasta de fotos", pasta.getUrl()); put("Recebido em", new Date()); put("dados_saida", JSON.stringify(dados));
      sh.appendRow(L);
      fixarTexto(sh, sh.getLastRow(), L);
    } finally { lock.releaseLock(); }
    return { ok: true };
  }
  // retorno
  let n = linhaDoId(sh, r.id);
  if (!n && r.placa) n = linhaAbertaDaPlaca(sh, r.placa);
  let saida = {};
  if (n) {
    if (sh.getRange(n, c("Data retorno")).getValue()) return { ok: true, repetido: true };
    try { saida = JSON.parse(sh.getRange(n, c("dados_saida")).getValue() || "{}"); } catch (e) { }
  }
  const pasta = pastaDoChecklist(r, saida.pastaId);
  const f = salvarFotos(pasta, r, "retorno");
  const fr = r.frio;
  if (fr && saida.frio) { fr.exigMin = saida.frio.exigMin; fr.exigMax = saida.frio.exigMax; }
  const dados = { km: r.km, clientesRet: r.clientesRet, pesoRet: r.pesoRet, obs: r.obs,
    frio: fr ? { temp: fr.temp, aparelho: fr.aparelho, exigMin: fr.exigMin, exigMax: fr.exigMax, fora: fora(fr.temp, fr) } : null,
    fotosIds: f.ids, avariasIds: f.av, assinaturaId: f.ass, motorista: r.motorista, criadoEm: r.criadoEm };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if (!n) n = linhaDoId(sh, r.id) || linhaAbertaDaPlaca(sh, r.placa);
    if (n && sh.getRange(n, c("Data retorno")).getValue()) return { ok: true, repetido: true };
    if (!n) {
      const L = new Array(COLS.length).fill("");
      L[c("ID") - 1] = r.id; L[c("Situação") - 1] = "Retorno sem saída"; L[c("Placa") - 1] = r.placa;
      L[c("Motorista") - 1] = txt(r.motorista && r.motorista.nome, 80); L[c("CPF") - 1] = txt(r.motorista && r.motorista.cpf, 14);
      L[c("Pasta de fotos") - 1] = pasta.getUrl(); L[c("Recebido em") - 1] = new Date();
      sh.appendRow(L); n = sh.getLastRow(); fixarTexto(sh, n, L);
    } else sh.getRange(n, c("Situação")).setValue("Aguardando validação");
    const kmS = Number(sh.getRange(n, c("Km saída")).getValue());
    const vals = {
      "Data retorno": dataDe(r.data), "Hora retorno": "'" + String(r.hora || "").slice(0, 5), "Km chegada": numero(r.km),
      "Km rodado": (r.km != null && kmS && r.km >= kmS) ? r.km - kmS : "", "Clientes retornados": numero(r.clientesRet), "Peso retornado": txt(r.pesoRet, 20),
      "Temp. retorno": txt(tempTxt(fr)), "Aparelho frio retorno": fr ? (fr.aparelho === "nok" ? "COM PROBLEMA" : "OK") : "",
      "Avarias retorno": txt(f.av.map((a, i) => (i + 1) + ". " + a.desc).join("; "), 2000), "Obs. retorno": txt(r.obs, 2000),
      "dados_retorno": JSON.stringify(dados)
    };
    Object.keys(vals).forEach(k => sh.getRange(n, c(k)).setValue(vals[k]));
  } finally { lock.releaseLock(); }
  return { ok: true };
}

/* ---------- administrativo ---------- */
function hmac(s) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, props().getProperty("SEGREDO"))); }
function assinarToken(o) { const s = Utilities.base64EncodeWebSafe(JSON.stringify(o)); return s + "." + hmac(s); }
function verificar(t) {
  const p = String(t || "").split(".");
  if (p.length !== 2 || hmac(p[0]) !== p[1]) throw new Error("login");
  const o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString());
  if (!o.exp || o.exp < Date.now()) throw new Error("login");
  return o;
}
function login(q) {
  const nome = limpo(String(q.nome || "").trim(), 60);
  if (!nome) return { ok: false, erro: "Informe seu nome" };
  let papel = null;
  if (q.senha && q.senha === SENHA_ADMINISTRATIVO) papel = "adm";
  else if (q.senha && q.senha === SENHA_DIRETORIA) papel = "dir";
  if (!papel || /^TROQUE/.test(q.senha)) { Utilities.sleep(1200); return { ok: false, erro: papel ? "Troque a senha padrão no Code.gs antes de usar" : "Senha incorreta" }; }
  return { ok: true, token: assinarToken({ p: papel, n: nome, exp: Date.now() + 12 * 3600e3 }), papel, nome };
}
function listar(q) {
  verificar(q.token);
  const sh = aba(), n = sh.getLastRow() - 1;
  const out = [];
  if (n > 0) {
    const V = sh.getRange(2, 1, n, COLS.length).getValues();
    const de = String(q.de || ""), ate = String(q.ate || "");
    for (let i = V.length - 1; i >= 0 && out.length < 1500; i--) {
      const v = V[i], g = k => v[c(k) - 1];
      const ds = iso(g("Data saída")) || iso(g("Data retorno")) || iso(g("Recebido em"));
      if (de && ds && ds < de) continue;
      if (ate && ds && ds > ate) continue;
      let s = null, r = null;
      try { s = g("dados_saida") ? JSON.parse(g("dados_saida")) : null; } catch (e) { }
      try { r = g("dados_retorno") ? JSON.parse(g("dados_retorno")) : null; } catch (e) { }
      out.push({ id: String(g("ID")), situacao: g("Situação"), dataSaida: iso(g("Data saída")), horaSaida: horaTxt(g("Hora saída")),
        placa: String(g("Placa")), motorista: String(g("Motorista")), cpf: String(g("CPF")), ajudante: String(g("Ajudante")), ajudante2: String(g("Ajudante 2")),
        dataRetorno: iso(g("Data retorno")), horaRetorno: horaTxt(g("Hora retorno")), kmRodado: g("Km rodado"), pasta: String(g("Pasta de fotos")),
        saida: s, retorno: r,
        validacao: g("Validado por") ? { situacao: g("Situação"), por: String(g("Validado por")), em: g("Validado em") instanceof Date ? Utilities.formatDate(g("Validado em"), FUSO, "dd/MM/yyyy HH:mm") : String(g("Validado em")), comentario: String(g("Comentário validação")) } : null });
    }
  }
  return { ok: true, linhas: out, planilha: SpreadsheetApp.openById(props().getProperty("PLANILHA_ID")).getUrl() };
}
function foto(q) {
  verificar(q.token);
  const f = DriveApp.getFileById(String(q.id));
  if (f.getName().indexOf("KP_") !== 0) throw new Error("arquivo");
  const b = f.getBlob();
  return { ok: true, tipo: b.getContentType(), b64: Utilities.base64Encode(b.getBytes()) };
}
function validar(q) {
  const u = verificar(q.token);
  if (u.p !== "adm") return { ok: false, erro: "Só o administrativo pode validar" };
  if (q.situacao !== "Validado" && q.situacao !== "Reprovado") return { ok: false, erro: "situação" };
  const sh = aba(), n = linhaDoId(sh, q.id);
  if (!n) return { ok: false, erro: "Check-list não encontrado" };
  const agora = new Date();
  sh.getRange(n, c("Situação")).setValue(q.situacao);
  sh.getRange(n, c("Validado por")).setValue(txt(u.n, 60));
  sh.getRange(n, c("Validado em")).setValue(agora);
  sh.getRange(n, c("Comentário validação")).setValue(txt(q.comentario, 1000));
  return { ok: true, validacao: { situacao: q.situacao, por: u.n, em: Utilities.formatDate(agora, FUSO, "dd/MM/yyyy HH:mm"), comentario: String(q.comentario || "") } };
}
