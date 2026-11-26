/* Fila offline do KP CHECKLIST.
   Usada pela página e pelo service worker: guarda os check-lists no celular
   (IndexedDB) e envia ao servidor quando houver internet. */
self.KPFila = (() => {
  const NOME = "kp-checklist", VER = 1;
  let conexao = null;
  function abrir() {
    if (conexao) return conexao;
    conexao = new Promise((ok, no) => {
      const r = indexedDB.open(NOME, VER);
      r.onupgradeneeded = () => {
        const db = r.result;
        if (!db.objectStoreNames.contains("fila")) db.createObjectStore("fila", { keyPath: "chave" });
        if (!db.objectStoreNames.contains("rascunho")) db.createObjectStore("rascunho", { keyPath: "chave" });
      };
      r.onsuccess = () => { const db = r.result; db.onclose = () => { conexao = null; }; ok(db); };
      r.onerror = () => { conexao = null; no(r.error); };
    });
    return conexao;
  }
  const req = r => new Promise((ok, no) => { r.onsuccess = () => ok(r.result); r.onerror = () => no(r.error); });
  async function loja(nome, modo) { const db = await abrir(); return db.transaction(nome, modo).objectStore(nome); }

  async function guardar(reg) { return req((await loja("fila", "readwrite")).put(reg)); }
  async function lista() {
    const t = await req((await loja("fila", "readonly")).getAll());
    return t.sort((a, b) => String(a.chave).localeCompare(String(b.chave)));
  }
  async function remover(chave) { return req((await loja("fila", "readwrite")).delete(chave)); }
  async function salvarRascunho(r) { return req((await loja("rascunho", "readwrite")).put({ chave: "atual", r })); }
  async function lerRascunho() { const x = await req((await loja("rascunho", "readonly")).get("atual")); return x ? x.r : null; }
  async function apagarRascunho() { return req((await loja("rascunho", "readwrite")).delete("atual")); }

  function comTempo(p, ms) {
    return Promise.race([p, new Promise((_, no) => setTimeout(() => no(new Error("tempo")), ms))]);
  }

  let rodando = null;
  function enviarTudo() {
    if (rodando) return rodando;
    rodando = (async () => {
      const C = self.KP_CONFIG || {};
      let enviados = 0, erro = null;
      if (!C.API_URL) return { enviados, pendentes: (await lista()).length, erro: "sem-config" };
      for (const r of await lista()) {
        try {
          const resp = await comTempo(fetch(C.API_URL, {
            method: "POST",
            body: JSON.stringify({ acao: "enviar", codigo: C.CODIGO, registro: r.dados })
          }), 120000);
          const j = await resp.json();
          if (j.ok) { await remover(r.chave); enviados++; }
          else {
            erro = j.erro || "erro";
            r.tentativas = (r.tentativas || 0) + 1; r.ultimoErro = erro;
            await guardar(r);
            if (erro === "codigo") break;
          }
        } catch (e) { erro = "rede"; break; }
      }
      return { enviados, pendentes: (await lista()).length, erro };
    })().finally(() => { rodando = null; });
    return rodando;
  }

  return { guardar, lista, remover, salvarRascunho, lerRascunho, apagarRascunho, enviarTudo };
})();
