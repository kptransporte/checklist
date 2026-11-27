/**
 * KP CHECKLIST — servidor (Google Apps Script) — versão 2
 * Recebe os check-lists do app, grava na planilha e guarda as fotos no Drive.
 *
 * ANTES DE USAR: troque a senha do Dev abaixo, salve e rode a função "configurar".
 * As outras pessoas (Administrativo e Diretoria) o Dev cadastra pelo próprio app.
 */
const NOME_DEV = "Patrick";
const SENHA_DEV_INICIAL = "TROQUE-senha-do-Dev";
const CODIGO_EMPRESA = "KP-7Q4M-2026"; // igual ao CODIGO do arquivo config.js do app

const ABA = "Checklists", ABA_USU = "Usuários", ABA_REG = "Registro de acessos", ABA_LIX = "Lixeira";
const PERFIS = { dev: "Dev", adm: "Administrativo", dir: "Diretoria" };
const USU_COLS = ["ID", "Nome", "Perfil", "Ativo", "Criado em", "Criado por", "Último acesso", "Senha alterada em", "hash", "sal"];
const FUSO = "America/Sao_Paulo";
const COLS = ["ID", "Situação", "Data saída", "Hora saída", "Placa", "Motorista", "CPF", "Ajudante",
  "Km saída", "Nº transporte", "Lacre", "Qtd. entregas", "Peso carga", "Itens com problema (NOK)",
  "Baú refrigerado", "Temp. exigida", "Temp. saída", "Aparelho frio saída", "Avarias saída", "Obs. saída",
  "Data retorno", "Hora retorno", "Km chegada", "Km rodado", "Entregas que voltaram", "Peso que voltou", "Divergência na rota",
  "Temp. retorno", "Aparelho frio retorno", "Avarias retorno", "Obs. retorno", "Pasta de fotos",
  "Validado por", "Validado em", "Comentário validação", "Recebido em", "dados_saida", "dados_retorno", "Anomalias na carga"];
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
    R().whenCellNotEmpty().setBackground("#fdecea").setFontColor("#c0150e").setBold(true).setRanges([tudo("Anomalias na carga")]).build(),
    R().whenTextStartsWith("Sim").setBackground("#fdecea").setFontColor("#c0150e").setBold(true)
      .setRanges([tudo("Divergência na rota")]).build(),
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
  // abas novas (v2)
  const cab = (nome, cols, cor) => {
    let x = ss.getSheetByName(nome); if (!x) x = ss.insertSheet(nome);
    x.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight("bold").setBackground(cor).setFontColor("#ffffff");
    x.setFrozenRows(1); return x;
  };
  const usu = cab(ABA_USU, USU_COLS, "#111418");
  usu.hideColumns(9, 2); usu.getRange(2, 1, usu.getMaxRows() - 1, 1).setNumberFormat("@");
  cab(ABA_REG, ["Data e hora", "Nome", "Perfil", "Ação", "Detalhe"], "#5b6370").getRange("A:A").setNumberFormat("dd/mm/yyyy hh:mm:ss");
  const lix = cab(ABA_LIX, COLS.concat(["Apagado por", "Apagado em", "Motivo"]), "#e6261f");
  lix.hideColumns(c("dados_saida"), 2);
  if (!lerUsuarios().some(u => u.perfil === "dev")) {
    if (/^TROQUE/.test(SENHA_DEV_INICIAL) || SENHA_DEV_INICIAL.length < 6) throw new Error("Troque a SENHA_DEV_INICIAL (linha 8) por uma senha de pelo menos 6 caracteres, salve e rode de novo.");
    novoUsuario(NOME_DEV, "dev", SENHA_DEV_INICIAL, "configuração");
  }
  Logger.log("Pronto! Planilha configurada. Se o app já estava publicado: Implantar > Gerenciar implantações > editar > Nova versão.");
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
      case "apagar": return json(apagar(q));
      case "usuarios": return json(usuarios(q));
      case "usuario_salvar": return json(usuarioSalvar(q));
      case "registro": return json(registro(q));
      case "minha_senha": return json(minhaSenha(q));
      case "sair": { const u = verificar(q.token); registrar(u.n, PERFIS[u.p], "Saiu", ""); return json({ ok: true }); }
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
    const anom = a.tipo === "anomalia" || /^ANOMALIA NA CARGA:/.test(String(a.desc || ""));
    const d = limpo(String(a.desc || "").replace(/^ANOMALIA NA CARGA:\s*/, ""), 300);
    av.push({ desc: anom ? "ANOMALIA NA CARGA: " + d : d, tipo: anom ? "anomalia" : "", foto: salvarImagem(pasta, a.foto, pl + "_" + fase + (anom ? "_anomalia" : "_avaria") + (i + 1) + ".jpg") });
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
      put("Ajudante", txt(r.ajudante, 80)); put("Km saída", numero(r.km));
      put("Nº transporte", txt(r.transporte, 40)); put("Lacre", txt(r.lacre, 40)); put("Qtd. entregas", numero(r.clientes)); put("Peso carga", txt(r.peso, 20));
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
  const dv = r.divergencia || {};
  const dados = { km: r.km, clientesRet: r.clientesRet, pesoRet: r.pesoRet, divergencia: { houve: !!dv.houve, desc: limpo(dv.desc, 1000) }, obs: r.obs,
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
      "Km rodado": (r.km != null && kmS && r.km >= kmS) ? r.km - kmS : "", "Entregas que voltaram": numero(r.clientesRet), "Peso que voltou": txt(r.pesoRet, 20),
      "Divergência na rota": dv.houve ? txt("Sim: " + (dv.desc || ""), 1000) : "Não",
      "Temp. retorno": txt(tempTxt(fr)), "Aparelho frio retorno": fr ? (fr.aparelho === "nok" ? "COM PROBLEMA" : "OK") : "",
      "Avarias retorno": txt(f.av.filter(a => !a.tipo).map((a, i) => (i + 1) + ". " + a.desc).join("; "), 2000),
      "Anomalias na carga": txt(f.av.filter(a => a.tipo).map((a, i) => (i + 1) + ". " + a.desc.replace(/^ANOMALIA NA CARGA:\s*/, "")).join("; "), 2000), "Obs. retorno": txt(r.obs, 2000),
      "dados_retorno": JSON.stringify(dados)
    };
    Object.keys(vals).forEach(k => sh.getRange(n, c(k)).setValue(vals[k]));
  } finally { lock.releaseLock(); }
  return { ok: true };
}

/* ---------- esqueceu a senha do Dev? ----------
 * Coloque a senha nova na linha SENHA_DEV_INICIAL (lá em cima), salve (Ctrl+S),
 * escolha "redefinirSenhaDev" na caixinha e clique em Executar.
 * Só quem tem acesso a este editor consegue fazer isso. Não precisa implantar de novo. */
function redefinirSenhaDev() {
  if (/^TROQUE/.test(SENHA_DEV_INICIAL) || SENHA_DEV_INICIAL.length < 6) throw new Error("Coloque a senha nova (mínimo 6 caracteres) na linha SENHA_DEV_INICIAL, salve e rode de novo.");
  const L = lerUsuarios();
  let x = L.find(u => norm(u.nome) === norm(NOME_DEV)) || L.find(u => u.perfil === "dev");
  if (!x) { novoUsuario(NOME_DEV, "dev", SENHA_DEV_INICIAL, "redefinição"); Logger.log("Usuário Dev criado: " + NOME_DEV); return; }
  const sh = usuSheet();
  sh.getRange(x.linha, 3).setValue(PERFIS.dev);
  sh.getRange(x.linha, 4).setValue("Sim");
  trocarSenha(x, SENHA_DEV_INICIAL);
  registrar(x.nome, "Dev", "Senha do Dev redefinida pelo editor", "");
  Logger.log("Pronto! Entre no app com o nome \"" + x.nome + "\" e a senha da linha SENHA_DEV_INICIAL.");
}

/* ---------- usuários e acessos ---------- */
function ss_() { return SpreadsheetApp.openById(props().getProperty("PLANILHA_ID")); }
function usuSheet() { return ss_().getSheetByName(ABA_USU); }
function codPerfil(t) { for (const k in PERFIS) if (PERFIS[k] === t) return k; return ""; }
function norm(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").trim().toLowerCase().replace(/\s+/g, " "); }
function hashSenha(senha, sal) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + "|" + senha, Utilities.Charset.UTF_8));
}
function lerUsuarios() {
  const sh = usuSheet(); if (!sh) return [];
  const n = sh.getLastRow() - 1; if (n < 1) return [];
  return sh.getRange(2, 1, n, USU_COLS.length).getValues().map((v, i) => ({
    linha: i + 2, id: String(v[0]), nome: String(v[1]), perfil: codPerfil(v[2]), ativo: v[3] === "Sim",
    criado: v[4], ultimo: v[6], senhaEm: v[7] instanceof Date ? v[7].getTime() : 0, hash: String(v[8]), sal: String(v[9])
  })).filter(u => u.id);
}
function novoUsuario(nome, perfil, senha, por) {
  const sal = Utilities.getUuid(), agora = new Date();
  usuSheet().appendRow(["'" + Utilities.getUuid().slice(0, 8), txt(nome, 60), PERFIS[perfil], "Sim", agora, txt(por, 60), "", agora, hashSenha(senha, sal), sal]);
}
function registrar(nome, perfil, acao, detalhe) {
  try { ss_().getSheetByName(ABA_REG).appendRow([new Date(), txt(nome, 60), txt(perfil, 30), txt(acao, 60), txt(detalhe, 300)]); } catch (e) { }
}
function hmac(s) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, props().getProperty("SEGREDO"))); }
function assinarToken(o) { const s = Utilities.base64EncodeWebSafe(JSON.stringify(o)); return s + "." + hmac(s); }
function verificar(t) {
  const p = String(t || "").split(".");
  if (p.length !== 2 || hmac(p[0]) !== p[1]) throw new Error("login");
  const o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString());
  const u = lerUsuarios().find(x => x.id === o.u);
  if (!u || !u.ativo || !u.perfil || (o.t || 0) < u.senhaEm) throw new Error("login");
  return { id: u.id, n: u.nome, p: u.perfil, linha: u.linha };
}
function login(q) {
  const nome = norm(q.nome), senha = String(q.senha || "");
  if (!nome || !senha) return { ok: false, erro: "Informe seu nome e sua senha" };
  const u = lerUsuarios().find(x => norm(x.nome) === nome);
  if (!u || !u.ativo || hashSenha(senha, u.sal) !== u.hash) {
    Utilities.sleep(1200); registrar(String(q.nome || "").slice(0, 60), "", "Login recusado", "");
    return { ok: false, erro: "Nome ou senha incorretos" };
  }
  usuSheet().getRange(u.linha, 7).setValue(new Date());
  registrar(u.nome, PERFIS[u.perfil], "Entrou", String(q.aparelho || "").slice(0, 120));
  return { ok: true, token: assinarToken({ u: u.id, t: Date.now() }), papel: u.perfil, nome: u.nome };
}
function minhaSenha(q) {
  const u = verificar(q.token), x = lerUsuarios().find(y => y.id === u.id);
  if (hashSenha(String(q.atual || ""), x.sal) !== x.hash) return { ok: false, erro: "Senha atual incorreta" };
  if (String(q.nova || "").length < 6) return { ok: false, erro: "A nova senha precisa ter pelo menos 6 caracteres" };
  trocarSenha(x, q.nova);
  registrar(u.n, PERFIS[u.p], "Trocou a própria senha", "");
  return { ok: true, token: assinarToken({ u: u.id, t: Date.now() + 1000 }) };
}
function trocarSenha(x, senha) {
  const sal = Utilities.getUuid(), sh = usuSheet();
  sh.getRange(x.linha, 8).setValue(new Date());
  sh.getRange(x.linha, 9, 1, 2).setValues([[hashSenha(String(senha), sal), sal]]);
}
function usuarios(q) {
  const u = verificar(q.token); if (u.p !== "dev") return { ok: false, erro: "Só o Dev gerencia usuários" };
  const fmt = d => d instanceof Date ? Utilities.formatDate(d, FUSO, "dd/MM/yyyy HH:mm") : "";
  return { ok: true, usuarios: lerUsuarios().map(x => ({ id: x.id, nome: x.nome, perfil: x.perfil, ativo: x.ativo, ultimo: fmt(x.ultimo), eu: x.id === u.id })) };
}
function usuarioSalvar(q) {
  const u = verificar(q.token); if (u.p !== "dev") return { ok: false, erro: "Só o Dev gerencia usuários" };
  const L = lerUsuarios(), nome = String(q.nome || "").trim().replace(/\s+/g, " ");
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (!q.id) {
      if (nome.length < 3) return { ok: false, erro: "Escreva o nome da pessoa" };
      if (!PERFIS[q.perfil]) return { ok: false, erro: "Escolha o perfil" };
      if (String(q.senha || "").length < 6) return { ok: false, erro: "A senha precisa ter pelo menos 6 caracteres" };
      if (L.some(x => norm(x.nome) === norm(nome))) return { ok: false, erro: "Já existe alguém com esse nome. Use nome e sobrenome." };
      novoUsuario(nome, q.perfil, q.senha, u.n);
      registrar(u.n, PERFIS[u.p], "Cadastrou usuário", nome + " (" + PERFIS[q.perfil] + ")");
      return { ok: true };
    }
    const x = L.find(y => y.id === q.id); if (!x) return { ok: false, erro: "Usuário não encontrado" };
    const sh = usuSheet();
    if (q.ativo !== undefined) {
      if (x.id === u.id) return { ok: false, erro: "Você não pode desativar a si mesmo" };
      sh.getRange(x.linha, 4).setValue(q.ativo ? "Sim" : "Não");
      registrar(u.n, PERFIS[u.p], q.ativo ? "Reativou usuário" : "Desativou usuário", x.nome);
    }
    if (q.perfil && PERFIS[q.perfil] && q.perfil !== x.perfil) {
      if (x.id === u.id) return { ok: false, erro: "Você não pode mudar o próprio perfil" };
      sh.getRange(x.linha, 3).setValue(PERFIS[q.perfil]);
      registrar(u.n, PERFIS[u.p], "Mudou perfil", x.nome + " → " + PERFIS[q.perfil]);
    }
    if (q.senha) {
      if (String(q.senha).length < 6) return { ok: false, erro: "A senha precisa ter pelo menos 6 caracteres" };
      trocarSenha(x, q.senha);
      registrar(u.n, PERFIS[u.p], "Definiu nova senha", x.nome);
    }
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function registro(q) {
  const u = verificar(q.token); if (u.p !== "dev") return { ok: false, erro: "Só o Dev vê o registro" };
  const sh = ss_().getSheetByName(ABA_REG), n = sh.getLastRow() - 1;
  if (n < 1) return { ok: true, linhas: [] };
  const k = Math.min(n, 400), V = sh.getRange(sh.getLastRow() - k + 1, 1, k, 5).getValues().reverse();
  return { ok: true, linhas: V.map(v => [v[0] instanceof Date ? Utilities.formatDate(v[0], FUSO, "dd/MM/yyyy HH:mm") : String(v[0]), String(v[1]), String(v[2]), String(v[3]), String(v[4])]) };
}

/* ---------- check-lists (escritório) ---------- */
function listar(q) {
  verificar(q.token);
  const sh = aba(), n = sh.getLastRow() - 1;
  const out = [];
  if (n > 0) {
    const V = sh.getRange(2, 1, n, COLS.length).getValues();
    const de = String(q.de || ""), ate = String(q.ate || "");
    for (let i = V.length - 1; i >= 0 && out.length < 1500; i--) {
      const v = V[i], g = k => v[c(k) - 1];
      if (!g("ID")) continue;
      const ds = iso(g("Data saída")) || iso(g("Data retorno")) || iso(g("Recebido em"));
      if (de && ds && ds < de) continue;
      if (ate && ds && ds > ate) continue;
      let s = null, r = null;
      try { s = g("dados_saida") ? JSON.parse(g("dados_saida")) : null; } catch (e) { }
      try { r = g("dados_retorno") ? JSON.parse(g("dados_retorno")) : null; } catch (e) { }
      out.push({ id: String(g("ID")), situacao: g("Situação"), dataSaida: iso(g("Data saída")), horaSaida: horaTxt(g("Hora saída")),
        placa: String(g("Placa")), motorista: String(g("Motorista")), cpf: String(g("CPF")), ajudante: String(g("Ajudante")),
        dataRetorno: iso(g("Data retorno")), horaRetorno: horaTxt(g("Hora retorno")), kmRodado: g("Km rodado"), pasta: String(g("Pasta de fotos")),
        saida: s, retorno: r,
        validacao: g("Validado por") ? { situacao: g("Situação"), por: String(g("Validado por")), em: g("Validado em") instanceof Date ? Utilities.formatDate(g("Validado em"), FUSO, "dd/MM/yyyy HH:mm") : String(g("Validado em")), comentario: String(g("Comentário validação")) } : null });
    }
  }
  return { ok: true, linhas: out, planilha: ss_().getUrl() };
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
  if (u.p !== "adm" && u.p !== "dev") return { ok: false, erro: "Só o Administrativo pode validar" };
  if (q.situacao !== "Validado" && q.situacao !== "Reprovado") return { ok: false, erro: "situação" };
  const sh = aba(), n = linhaDoId(sh, q.id);
  if (!n) return { ok: false, erro: "Check-list não encontrado" };
  const agora = new Date();
  sh.getRange(n, c("Situação")).setValue(q.situacao);
  sh.getRange(n, c("Validado por")).setValue(txt(u.n, 60));
  sh.getRange(n, c("Validado em")).setValue(agora);
  sh.getRange(n, c("Comentário validação")).setValue(txt(q.comentario, 1000));
  registrar(u.n, PERFIS[u.p], q.situacao === "Validado" ? "Validou check-list" : "Reprovou check-list", sh.getRange(n, c("Placa")).getValue() + " · " + iso(sh.getRange(n, c("Data saída")).getValue()));
  return { ok: true, validacao: { situacao: q.situacao, por: u.n, em: Utilities.formatDate(agora, FUSO, "dd/MM/yyyy HH:mm"), comentario: String(q.comentario || "") } };
}
function apagar(q) {
  const u = verificar(q.token);
  if (u.p !== "adm" && u.p !== "dev") return { ok: false, erro: "Só o Administrativo e o Dev podem apagar" };
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const sh = aba(), n = linhaDoId(sh, q.id);
    if (!n) return { ok: false, erro: "Check-list não encontrado (talvez já apagado)" };
    const v = sh.getRange(n, 1, 1, COLS.length).getValues()[0];
    ss_().getSheetByName(ABA_LIX).appendRow(v.concat([txt(u.n, 60), new Date(), txt(q.motivo, 300)]));
    let pastaId = "";
    try { pastaId = JSON.parse(v[c("dados_saida") - 1] || "{}").pastaId || ""; } catch (e) { }
    if (!pastaId) { const m = /folders\/([A-Za-z0-9_-]+)/.exec(String(v[c("Pasta de fotos") - 1])); if (m) pastaId = m[1]; }
    if (pastaId) { try { DriveApp.getFolderById(pastaId).setTrashed(true); } catch (e) { } }
    sh.deleteRow(n);
    registrar(u.n, PERFIS[u.p], "Apagou check-list", v[c("Placa") - 1] + " · " + iso(v[c("Data saída") - 1]) + " · " + v[c("Motorista") - 1] + (q.motivo ? " · motivo: " + String(q.motivo).slice(0, 150) : ""));
    return { ok: true };
  } finally { lock.releaseLock(); }
}
