/* KP CHECKLIST — gera arquivo Excel (.xlsx) no próprio aparelho, sem internet e sem conta Google.
   Uso: KPXlsx.gerar([{nome, colunas:[{t, w}], linhas:[[valor|{v, s}]]}]) -> Blob
   Valor: texto, número, Date (vira data do Excel) ou {v, s} com estilo: "nok" (vermelho), "ok" (verde). */
(function () {
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const enc = s => new TextEncoder().encode(s);

  // ZIP simples (sem compressão), suficiente para o Excel abrir
  function zip(arquivos) {
    const partes = [], centrais = []; let pos = 0;
    for (const [nome, txt] of arquivos) {
      const n = enc(nome), d = enc(txt), crc = crc32(d);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true);
      h.setUint32(14, crc, true); h.setUint32(18, d.length, true); h.setUint32(22, d.length, true); h.setUint16(26, n.length, true);
      partes.push(new Uint8Array(h.buffer), n, d);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint32(16, crc, true); c.setUint32(20, d.length, true); c.setUint32(24, d.length, true); c.setUint16(28, n.length, true); c.setUint32(42, pos, true);
      centrais.push(new Uint8Array(c.buffer), n);
      pos += 30 + n.length + d.length;
    }
    const tam = centrais.reduce((s, x) => s + x.length, 0);
    const f = new DataView(new ArrayBuffer(22));
    f.setUint32(0, 0x06054b50, true); f.setUint16(8, arquivos.length, true); f.setUint16(10, arquivos.length, true);
    f.setUint32(12, tam, true); f.setUint32(16, pos, true);
    return new Blob([...partes, ...centrais, new Uint8Array(f.buffer)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  const x = v => String(v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const col = i => { let s = ""; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  // estilos: 0 normal, 1 cabeçalho, 2 data, 3 NOK (vermelho), 4 OK (verde), 5 texto com quebra
  const EST = { nok: 3, ok: 4 };
  const serial = d => (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes()) - Date.UTC(1899, 11, 30)) / 864e5;

  function planilha(p) {
    const C = p.colunas, ult = col(C.length - 1);
    let s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>';
    C.forEach((c, i) => { s += '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + (c.w || 14) + '" customWidth="1"/>'; });
    s += '</cols><sheetData><row r="1">';
    C.forEach((c, i) => { s += '<c r="' + col(i) + '1" s="1" t="inlineStr"><is><t>' + x(c.t) + '</t></is></c>'; });
    s += '</row>';
    p.linhas.forEach((L, j) => {
      const r = j + 2; s += '<row r="' + r + '">';
      L.forEach((cel, i) => {
        let v = cel, st = 0;
        if (cel && typeof cel === "object" && !(cel instanceof Date)) { v = cel.v; st = EST[cel.s] || 0; }
        if (v == null || v === "") { if (st) s += '<c r="' + col(i) + r + '" s="' + st + '"/>'; return; }
        const ref = col(i) + r;
        if (v instanceof Date && !isNaN(v)) s += '<c r="' + ref + '" s="2"><v>' + serial(v) + '</v></c>';
        else if (typeof v === "number" && isFinite(v)) s += '<c r="' + ref + '"' + (st ? ' s="' + st + '"' : '') + '><v>' + v + '</v></c>';
        else s += '<c r="' + ref + '" s="' + (st || 5) + '" t="inlineStr"><is><t xml:space="preserve">' + x(v) + '</t></is></c>';
      });
      s += '</row>';
    });
    s += '</sheetData>';
    if (p.linhas.length) s += '<autoFilter ref="A1:' + ult + (p.linhas.length + 1) + '"/>';
    return s + '</worksheet>';
  }

  function gerar(abas) {
    const A = [];
    A.push(["[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + abas.map((_, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') + '</Types>']);
    A.push(["_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>']);
    A.push(["xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + abas.map((a, i) => '<sheet name="' + x(String(a.nome).replace(/[\\\/\?\*\[\]:]/g, "").slice(0, 31)) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join('') + '</sheets>' + (abas.some(a => a.linhas.length) ? '<definedNames>' + abas.map((a, i) => a.linhas.length ? '<definedName name="_xlnm._FilterDatabase" localSheetId="' + i + '" hidden="1">\'' + x(String(a.nome).replace(/[\\\/\?\*\[\]:]/g, "").slice(0, 31)) + '\'!$A$1:$' + col(a.colunas.length - 1) + '$' + (a.linhas.length + 1) + '</definedName>' : '').join('') + '</definedNames>' : '') + '</workbook>']);
    A.push(["xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + abas.map((_, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') + '<Relationship Id="rId' + (abas.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>']);
    A.push(["xl/styles.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts><fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFC00000"/><name val="Calibri"/></font><font><sz val="11"/><color rgb="FF1E7B34"/><name val="Calibri"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFC8102E"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFDE2E2"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>']);
    abas.forEach((a, i) => A.push(["xl/worksheets/sheet" + (i + 1) + ".xml", planilha(a)]));
    return zip(A);
  }
  self.KPXlsx = { gerar };
})();
