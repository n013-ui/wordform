// 專題寫作排版工具（網頁版）：把 Gemini 的「排版用文字」轉成 Word 檔（.docx）
// 版面與 Claude 版相同：A4、邊界 2 公分、標楷體 12 號、固定行高、凸排縮排、表格不跨頁。
// 這支程式在瀏覽器（使用 docx.iife.js）與 Node（測試用）都能執行。
(function (root) {
  const D = (typeof docx !== 'undefined') ? docx : require('docx');
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
    ShadingType, BorderStyle, LineRuleType, VerticalAlign, HeightRule } = D;

  const FONT = '標楷體', SZ = 24;               // 12 號字
  const PAGE_W_TW = 11906 - 2 * 1134;           // 9638 twips（A4 扣左右 2 公分）
  const PAGE_H_PT = (16838 - 2 * 1134) / 20;    // 728.5 點可用高度
  const IND = 480;                              // 一個字寬（twips）
  const TABLE_LH = 360;                         // 表格內固定行高 18 點
  const NAPKIN_CM = 4.5;

  // ---------------- 解析 ----------------
  const RE_H1 = /^[一二三四五六七八九十]+、/;
  const RE_L1 = /^[（(][一二三四五六七八九十]+[）)]/;
  const RE_L2 = /^\d+\s*[.．、]/;
  const RE_SUBQ = /^(\d+\s*[.．、]|[（(][一二三四五六七八九十\d]+[）)])/;
  const RE_BREAK = /^(（此處分頁）|\(此處分頁\)|={3,}|＝{3,})\s*$/;
  const RE_NAPKIN = /^【\s*Napkin|^【.*概念圖/i;
  const RE_CAPTION = /^圖[一二三四五六七八九十]+[\s　]/;

  function clean(s) {
    return s.replace(/\*\*/g, '').replace(/^\s*#+\s*/, '').replace(/^\s*[-•]\s+/, '').replace(/\s+$/, '');
  }
  function splitLabel(s) {
    const i = s.indexOf('：');
    return (i > 0 && i <= 25) ? { label: s.slice(0, i + 1), text: s.slice(i + 1) } : { label: '', text: s };
  }
  function parse(raw) {
    const lines = String(raw || '').replace(/\r/g, '').split('\n').map(clean);
    const qs = []; let q, inRefs = false, table = null;
    const flush = () => { if (table) { q.blocks.push({ type: 'table', rows: table }); table = null; } };
    const newQ = () => { q = { title: '', prompt: [], blocks: [] }; qs.push(q); inRefs = false; };
    newQ();
    for (const line of lines) {
      const t = line.trim();
      if (RE_BREAK.test(t)) { flush(); if (q.title || q.blocks.length) newQ(); continue; }
      if (/^```/.test(t)) continue;
      if (t.startsWith('|')) {
        if (/^\|?[\s:|-]+\|?$/.test(t) && t.includes('-')) continue;
        (table = table || []).push(t.replace(/^\||\|$/g, '').split('|').map(c => c.trim()));
        continue;
      }
      flush();
      if (!t) continue;
      if (!q.blocks.length && !RE_H1.test(t)) { if (!q.title) q.title = t; else q.prompt.push(t); continue; }
      if (RE_H1.test(t)) { inRefs = /參考/.test(t); q.blocks.push({ type: 'h1', text: t }); continue; }
      if (RE_NAPKIN.test(t)) { const m = t.match(/[：:]\s*(.+?)】?$/); q.blocks.push({ type: 'napkin', title: m ? m[1].replace(/】$/, '') : '' }); continue; }
      if (RE_CAPTION.test(t)) { q.blocks.push({ type: 'caption', text: t }); continue; }
      if (inRefs && (RE_L2.test(t) || /https?:\/\//.test(t))) { q.blocks.push({ type: 'ref', text: t }); continue; }
      let m = t.match(RE_L1);
      if (m) { q.blocks.push(Object.assign({ type: 'item', level: 1, num: m[0] }, splitLabel(t.slice(m[0].length).trim()))); continue; }
      m = t.match(RE_L2);
      if (m) { q.blocks.push(Object.assign({ type: 'item', level: 2, num: m[0].replace(/\s/g, '') + ' ' }, splitLabel(t.slice(m[0].length).trim()))); continue; }
      q.blocks.push({ type: 'body', text: t });
    }
    flush();
    return qs.filter(x => x.title);
  }

  // ---------------- 版面規格（Word 與估算共用） ----------------
  // 縮排單位：twips。left＝其餘行起點，first＝第一行起點
  function specs(q) {
    const out = [];
    out.push({ kind: 'p', runs: [[q.title, true]], left: 0, first: 0, align: 'left', keepNext: true });
    q.prompt.forEach((t, i) => {
      const sub = RE_SUBQ.test(t);
      out.push({ kind: 'p', runs: [[t, false]], left: sub ? 2 * IND : 0, first: sub ? IND : IND,
        after: i === q.prompt.length - 1 ? 80 : 0 });
    });
    for (const b of q.blocks) {
      if (b.type === 'h1') out.push({ kind: 'p', runs: [[b.text, true]], left: 0, first: 0, before: 80, keepNext: true, align: 'left' });
      else if (b.type === 'body') out.push({ kind: 'p', runs: [[b.text, false]], left: IND, first: 2 * IND });
      else if (b.type === 'caption') out.push({ kind: 'p', runs: [[b.text, false]], left: 0, first: 0, align: 'center' });
      else if (b.type === 'ref') out.push({ kind: 'p', runs: [[b.text, false]], left: 2.5 * IND, first: IND, align: 'left' });
      else if (b.type === 'item') {
        // （一）：左 4 字、凸排 3 字；1.：左 3.5 字、凸排 1.5 字（與 Claude 版 Word 相同）
        const lv = b.level === 1 ? { left: 4 * IND, first: IND } : { left: 3.5 * IND, first: 2 * IND };
        const runs = [[b.num, false]];
        if (b.label) runs.push([b.label, true]);
        if (b.text) runs.push([b.text, false]);
        out.push(Object.assign({ kind: 'p', runs, keepNext: !b.text }, lv));
      } else if (b.type === 'table') out.push({ kind: 'table', rows: b.rows });
      else if (b.type === 'napkin') out.push({ kind: 'napkin', title: b.title });
    }
    return out;
  }
  function tableWidths(rows) {
    const n = Math.max(...rows.map(r => r.length));
    const w = PAGE_W_TW - 200;
    const firstChars = Math.max(...rows.map(r => (r[0] || '').length));
    const first = Math.round(Math.min(w * 0.3, Math.max(1300, firstChars * 240 + 260)));
    const ws = [first];
    for (let j = 1; j < n; j++) ws.push(Math.round((w - first) / (n - 1)));
    return n === 1 ? [w] : ws;
  }

  // ---------------- 產生 Word ----------------
  function build(questions, opt) {
    const LH = (opt && opt.lineHeight) || 400;
    const LINE = { line: LH, lineRule: LineRuleType.EXACT };
    const run = (text, bold) => new TextRun({ text, size: SZ, bold,
      font: { name: FONT, eastAsia: FONT, ascii: 'Times New Roman', hAnsi: 'Times New Roman' } });
    const AL = { left: AlignmentType.LEFT, center: AlignmentType.CENTER };
    const border = { style: BorderStyle.SINGLE, size: 4, color: '666666' };
    const borders = { top: border, bottom: border, left: border, right: border };
    const children = [];
    questions.forEach((q, qi) => {
      specs(q).forEach((s, si) => {
        if (s.kind === 'p') {
          const hang = s.left - s.first;
          children.push(new Paragraph({
            children: s.runs.map(([t, b]) => run(t, b)),
            alignment: AL[s.align] || AlignmentType.JUSTIFIED,
            spacing: Object.assign({}, LINE, { before: s.before || 0, after: s.after || 0 }),
            indent: hang > 0 ? { left: s.left, hanging: hang } : { left: s.left, firstLine: -hang },
            keepNext: s.keepNext, pageBreakBefore: qi > 0 && si === 0,
          }));
        } else if (s.kind === 'table') {
          const widths = tableWidths(s.rows), n = widths.length, R = s.rows.length;
          children.push(new Table({
            width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA }, columnWidths: widths,
            alignment: AlignmentType.CENTER,
            rows: s.rows.map((r, i) => new TableRow({ cantSplit: true, tableHeader: i === 0,
              children: widths.map((w, j) => {
                const head = i === 0, firstCol = i > 0 && j === 0;
                return new TableCell({ width: { size: w, type: WidthType.DXA }, borders, verticalAlign: VerticalAlign.CENTER,
                  shading: head ? { fill: 'DCE8F5', type: ShadingType.CLEAR, color: 'auto' }
                    : firstCol ? { fill: 'F2F6FB', type: ShadingType.CLEAR, color: 'auto' } : undefined,
                  margins: { top: 20, bottom: 20, left: 80, right: 80 },
                  children: [new Paragraph({ children: [run(r[j] || '', head || firstCol)], keepNext: i < R - 1, keepLines: true,
                    alignment: head || firstCol ? AlignmentType.CENTER : AlignmentType.LEFT,
                    spacing: { line: Math.min(LH, TABLE_LH), lineRule: LineRuleType.EXACT } })] });
              }) })),
          }));
        } else if (s.kind === 'napkin') {
          const dash = { style: BorderStyle.DASHED, size: 6, color: '999999' };
          const W = PAGE_W_TW - 200;
          const p = t => new Paragraph({ children: [run(t, false)], alignment: AlignmentType.CENTER, spacing: LINE });
          children.push(new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], alignment: AlignmentType.CENTER,
            rows: [new TableRow({ cantSplit: true, height: { value: Math.round(NAPKIN_CM * 567), rule: HeightRule.EXACT },
              children: [new TableCell({ width: { size: W, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
                borders: { top: dash, bottom: dash, left: dash, right: dash },
                children: [p('【Napkin 概念圖預留位置】' + s.title), p('（依提示詞生成後，刪除此框並插入圖片）')] })] })] }));
        }
      });
    });
    return new Document({
      styles: { default: { document: { run: { font: FONT, size: SZ } } } },
      sections: [{ properties: { page: { size: { width: 11906, height: 16838 },
        margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } }, children }],
    });
  }

  // ---------------- 估算頁數（固定行高，所以可以準確估算） ----------------
  // 字寬（點）：中文與全形 12；英數依 Times New Roman 約略寬度
  function cw(ch) {
    const c = ch.codePointAt(0);
    if (c > 0x2E7F) return 12;                          // CJK 與全形符號
    if (/[A-Z]/.test(ch)) return 8;
    if (/[a-z0-9]/.test(ch)) return 5.6;
    if (ch === ' ') return 3;
    return 4.5;
  }
  const NO_START = '，。、；：？！」』）】》〉%,.;:?!)]';   // 不放行首的標點
  function countLines(text, firstW, restW) {
    let lines = 1, x = 0, w = firstW;
    const chars = Array.from(text);
    for (let i = 0; i < chars.length; i++) {
      const ch = chars[i], d = cw(ch);
      if (x + d > w + 0.01) {
        if (NO_START.includes(ch) && x + d <= w + 6.5) { x += d; continue; }  // 標點擠進行尾
        lines++; x = 0; w = restW;
      }
      x += d;
    }
    return lines;
  }
  function estimate(q, opt) {
    const lhPt = ((opt && opt.lineHeight) || 400) / 20;
    const textW = PAGE_W_TW / 20;
    const blocks = specs(q).map(s => {
      if (s.kind === 'p') {
        const text = s.runs.map(r => r[0]).join('');
        const n = countLines(text, textW - s.first / 20, textW - s.left / 20);
        return { h: n * lhPt + ((s.before || 0) + (s.after || 0)) / 20, lines: n, lh: lhPt, keepNext: s.keepNext, split: true };
      }
      if (s.kind === 'table') {
        const ws = tableWidths(s.rows);
        let h = 0;
        s.rows.forEach(r => {
          let rh = 0;
          ws.forEach((w, j) => { rh = Math.max(rh, countLines(r[j] || '', w / 20 - 8, w / 20 - 8) * TABLE_LH / 20); });
          h += rh + 2.5;
        });
        return { h, split: false };
      }
      return { h: NAPKIN_CM * 28.35 + 1, split: false };
    });
    // 排版模擬：段落可以跨頁，表格與概念圖框不跨頁，標題與下一段同頁
    let page = 1, y = 0;
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.keepNext && blocks[i + 1]) {
        const nx = blocks[i + 1];
        const need = b.h + (nx.split ? nx.lh : nx.h);
        if (y + need > PAGE_H_PT + 0.01 && y > 0) { page++; y = 0; }
      }
      if (!b.split) {
        if (y + b.h > PAGE_H_PT + 0.01 && y > 0) { page++; y = 0; }
        y += b.h;
      } else {
        const extra = b.h - b.lines * b.lh;   // 段前段後
        y += extra;
        for (let k = 0; k < b.lines; k++) {
          if (y + b.lh > PAGE_H_PT + 0.01) { page++; y = 0; }
          y += b.lh;
        }
      }
    }
    const free = Math.floor((PAGE_H_PT - y) / lhPt);
    const linesPerPage = Math.floor(PAGE_H_PT / lhPt);
    return { pages: page, free, linesPerPage, fill: y / PAGE_H_PT };
  }

  function verdict(e, n) {
    const tag = '題目' + '一二三四五'[n - 1];
    const over = e.linesPerPage - e.free;
    if (e.pages > 2) return { ok: false, head: `超過兩頁（第 ${e.pages} 頁約有 ${over} 行）`,
      say: `${tag}第三頁多了 ${over} 行，請精簡 ${over + 1} 行，用修改對照給我。` };
    if (e.pages < 2) return { ok: false, head: `只有一頁（還空約 ${e.free} 行）`,
      say: `${tag}只有一頁，第二頁整頁都空著，請補充新的事實與台灣的例子，寫滿兩頁。` };
    if (e.fill < 0.8) return { ok: false, head: `第二頁還空約 ${e.free} 行，太空了`,
      say: `${tag}第二頁還空 ${e.free} 行，請補充新的事實，用修改對照給我。` };
    return { ok: true, head: `剛好兩頁，第二頁約寫到 ${Math.round(e.fill * 100)}%（還空約 ${e.free} 行）`, say: '' };
  }

  const api = { parse, build, estimate, verdict, Packer };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.ZT = api;
})(this);
