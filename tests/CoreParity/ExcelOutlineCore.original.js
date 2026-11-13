
/* =========================================================================
 * Excel 分级显示核心（浏览器 / Node 通用）
 *  - 智能识别中文概算表的编号层级
 *  - 只修改 XLSX 内部的分级属性（outlineLevel / hidden / collapsed / summaryBelow）
 *  - 单元格内容、样式、合并、行高列宽一律不碰
 * ========================================================================= */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.ExcelOutlineCore = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  const MAX_LEVEL = 7;

  const CODE_RULES = [
    ['part', /^第[一二三四五六七八九十百零〇\d]{1,4}\s*(项|部分|章|节|篇)/],
    ['roman', /^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫⅬⅭⅮⅯ]{1,4}$/],
    ['cn_paren', /^[（(]\s*[一二三四五六七八九十百零〇]{1,3}\s*[)）]$/],
    ['cn', /^[一二三四五六七八九十百零〇]{1,4}\s*[、.．,，]?$/],
    ['num_paren', /^[（(]\s*\d{1,3}\s*[)）]$/],
    ['num_ju', /^\d{1,3}\s*[)）]$/],
    ['dotted', /^\d{1,2}(?:[.．]\d{1,2}){1,4}[.．]?$/],
    ['num', /^\d{1,3}\s*[、.．]?$/],
    ['circle', /^[①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]$/],
  ];

  const FAMILY_PARENTS = {
    roman: [],
    part: [],
    cn: ['roman', 'part'],
    cn_paren: ['cn', 'roman', 'part'],
    num: ['cn_paren', 'cn', 'roman', 'part'],
    num_paren: ['num', 'cn_paren', 'cn', 'roman', 'part'],
    num_ju: ['num_paren', 'num', 'cn_paren', 'cn', 'roman', 'part'],
    dotted: ['num', 'cn_paren', 'cn', 'roman', 'part'],
    circle: ['num_paren', 'num_ju', 'num', 'cn_paren', 'cn', 'roman', 'part'],
  };

  const FAMILY_LABELS = {
    part: '第X项 / 第X部分',
    roman: 'Ⅰ Ⅱ Ⅲ',
    cn: '一、二、',
    cn_paren: '（一）（二）',
    num: '1、2、',
    num_paren: '（1）（2）',
    num_ju: '1）2）',
    dotted: '3.1 / 3.1.2',
    circle: '① ②',
  };

  const HEADER_KEYS = ['编号', '序号', '项目编号', '编码'];
  const NAME_KEYS = ['名称', '项目名称', '名 称', '工程项目', '费用名称'];
  const SECTION_TOTAL_WORDS = ['一至', '静态投资', '基本预备费', '增值税', '总投资', '建设期利息'];

  let DOMParserImpl = typeof DOMParser !== 'undefined' ? DOMParser : null;

  function getParser() {
    if (!DOMParserImpl) {
      if (typeof require === 'function') {
        DOMParserImpl = require('@xmldom/xmldom').DOMParser;
      } else {
        throw new Error('当前环境不支持 XML 解析');
      }
    }
    return DOMParserImpl;
  }

  function parseXml(text) {
    return new (getParser())().parseFromString(text, 'text/xml');
  }

  function norm(value) {
    if (value === null || value === undefined) return '';
    return String(value).trim().replace(/\u3000/g, ' ').replace(/\u00a0/g, ' ');
  }

  // [v2] 数值型编号还原：Excel 会把 2.2 存成 2.2000000000000002、1.12 存成
  // 1.1200000000000001。若直接按文本匹配，编号规则（最多 2 位小数）认不出，
  // 整行会被当成「无编号明细」，导致层级塌陷、父子归属错乱。
  // 仅处理「整数 1~2 位 + 小数 >=3 位」且能无损还原成 1~2 位小数的串；
  // 金额（0.397 / 13.011 等）无法无损还原，会原样返回，不会被误判为编号。
  function decodeNumericCode(s) {
    if (!/^\d{1,2}\.\d{3,}$/.test(s)) return s;
    const n = Number(s);
    if (!isFinite(n)) return s;
    for (let k = 1; k <= 2; k++) {
      let t = n.toFixed(k).replace(/0+$/, '').replace(/\.$/, '');
      if (Math.abs(Number(t) - n) < 1e-8 && /^\d{1,2}(?:\.\d{1,2})?$/.test(t)) return t;
    }
    return s;
  }

  function codeFamily(value) {
    const s = decodeNumericCode(norm(value));
    if (!s || s.length > 40) return { family: null, depth: 0, text: s, parts: [] };
    for (let i = 0; i < CODE_RULES.length; i++) {
      const family = CODE_RULES[i][0];
      const rule = CODE_RULES[i][1];
      if (rule.test(s)) {
        if (family === 'dotted') {
          const depth = (s.match(/[.．]/g) || []).length;
          const parts = s.replace(/[.．]+$/, '').split(/[.．]/).map(function (x) {
            return parseInt(x, 10);
          });
          return { family: family, depth: Math.max(1, depth), text: s, parts: parts };
        }
        return { family: family, depth: 1, text: s, parts: [] };
      }
    }
    return { family: null, depth: 0, text: s, parts: [] };
  }

  const CN_DIGITS = { '零': 0, '〇': 0, '一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9 };
  const CN_UNITS = { '十': 10, '百': 100, '千': 1000 };
  const ROMAN_DIGITS = { 'Ⅰ': 1, 'Ⅱ': 2, 'Ⅲ': 3, 'Ⅳ': 4, 'Ⅴ': 5, 'Ⅵ': 6, 'Ⅶ': 7, 'Ⅷ': 8, 'Ⅸ': 9, 'Ⅹ': 10, 'Ⅺ': 11, 'Ⅻ': 12 };

  function cnToNumber(text) {
    const s = norm(text).replace(/[\s、.．,，()（）]/g, '');
    if (!s) return null;
    if (/^\d+$/.test(s)) return parseInt(s, 10);
    let total = 0;
    let section = 0;
    let number = null;
    for (let i = 0; i < s.length; i++) {
      const ch = s.charAt(i);
      if (CN_DIGITS[ch] !== undefined) {
        number = CN_DIGITS[ch];
      } else if (CN_UNITS[ch]) {
        if (number === null) number = 1;
        section += number * CN_UNITS[ch];
        number = null;
      } else {
        return null;
      }
    }
    return total + section + (number === null ? 0 : number);
  }

  // 解析编号自身的序号，用于点号编号（3.1 / 3.1.2）与上级编号按“首段数字”配对
  function codeOrdinal(family, text) {
    const s = norm(text);
    if (!s) return null;
    if (family === 'part') {
      const m = /^第\s*([一二三四五六七八九十百零〇\d]{1,4})\s*(?:项|部分|章|节|篇)/.exec(s);
      return m ? cnToNumber(m[1]) : null;
    }
    if (family === 'roman') {
      const m = /^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫ]{1,4}$/.exec(s);
      if (!m) return null;
      return ROMAN_DIGITS[s.charAt(0)] || null;
    }
    if (family === 'circle') {
      const circle = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳';
      const idx = circle.indexOf(s.charAt(0));
      return idx >= 0 ? idx + 1 : null;
    }
    if (family === 'cn' || family === 'cn_paren') return cnToNumber(s);
    if (family === 'num' || family === 'num_paren' || family === 'num_ju') {
      const m = /\d{1,3}/.exec(s);
      return m ? parseInt(m[0], 10) : null;
    }
    return null;
  }

  function familyLabel(family) {
    return FAMILY_LABELS[family] || family;
  }

  function attrOf(tag, name) {
    const re = new RegExp('(?:^|\\s)' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '="([^"]*)"');
    const m = re.exec(tag);
    return m ? m[1] : null;
  }

  function escapeXml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function colIndex(ref) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref || '');
    if (!m) return null;
    let n = 0;
    for (const ch of m[1]) n = n * 26 + (ch.charCodeAt(0) - 64);
    return { col: n, row: parseInt(m[2], 10) };
  }

  function colLetter(n) {
    let s = '';
    let x = n;
    while (x > 0) {
      const r = (x - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      x = Math.floor((x - 1) / 26);
    }
    return s;
  }

  // ---------------------------------------------------------------- 读取数据

  async function readSharedStrings(zip) {
    const file = zip.file('xl/sharedStrings.xml');
    if (!file) return [];
    const doc = parseXml(await file.async('string'));
    const sis = doc.getElementsByTagName('si');
    const out = [];
    for (let i = 0; i < sis.length; i++) {
      const ts = sis[i].getElementsByTagName('t');
      let text = '';
      for (let j = 0; j < ts.length; j++) {
        const parent = ts[j].parentNode;
        if (parent && parent.nodeName === 'rPh') continue;
        text += ts[j].textContent || '';
      }
      out.push(text);
    }
    return out;
  }

  function readSheetValues(xmlText, sharedStrings) {
    const doc = parseXml(xmlText);
    const rows = {};
    const formulas = {};
    let maxRow = 0;
    let maxCol = 0;
    const rowEls = doc.getElementsByTagName('row');
    for (let i = 0; i < rowEls.length; i++) {
      const rowEl = rowEls[i];
      const rAttr = rowEl.getAttribute('r');
      const rowNum = rAttr ? parseInt(rAttr, 10) : i + 1;
      if (!rowNum) continue;
      if (rowNum > maxRow) maxRow = rowNum;
      const cells = rowEl.getElementsByTagName('c');
      const row = rows[rowNum] || (rows[rowNum] = {});
      for (let j = 0; j < cells.length; j++) {
        const c = cells[j];
        const ref = c.getAttribute('r');
        const pos = ref ? colIndex(ref) : { col: j + 1, row: rowNum };
        if (!pos) continue;
        if (pos.col > maxCol) maxCol = pos.col;
        const t = c.getAttribute('t');
        let val = null;
        const fEl = c.getElementsByTagName('f')[0];
        if (fEl) {
          if (!formulas[rowNum]) formulas[rowNum] = {};
          formulas[rowNum][pos.col] = fEl.textContent || '';
        }
        if (t === 'inlineStr') {
          const is = c.getElementsByTagName('is')[0];
          if (is) {
            const ts = is.getElementsByTagName('t');
            let s = '';
            for (let k = 0; k < ts.length; k++) s += ts[k].textContent || '';
            val = s;
          }
        } else {
          const v = c.getElementsByTagName('v')[0];
          if (v) {
            const raw = v.textContent;
            if (t === 's') val = sharedStrings[parseInt(raw, 10)] || '';
            else val = raw;
          }
        }
        if (val !== null && val !== '') row[pos.col] = val;
      }
    }
    return { rows: rows, maxRow: maxRow, maxCol: maxCol, formulas: formulas };
  }

  function valAt(values, r, c) {
    const row = values[r];
    if (!row) return null;
    const v = row[c];
    return v === undefined ? null : v;
  }

  function rowIsEmpty(values, r, maxCol) {
    const row = values[r];
    if (!row) return true;
    for (let c = 1; c <= maxCol; c++) {
      const v = row[c];
      if (v !== undefined && v !== null && v !== '') return false;
    }
    return true;
  }

  function lastDataRow(values, maxRow, maxCol) {
    let last = 0;
    const cols = Math.max(1, maxCol || 0);
    for (let r = maxRow; r >= 1; r--) {
      if (!rowIsEmpty(values, r, cols)) { last = r; break; }
    }
    return last;
  }

  // ---------------------------------------------------------------- 表头识别

  function findHeader(values, maxRow, maxCol) {
    const scanRows = Math.min(maxRow, 30);
    let best = null;
    for (let r = 1; r <= scanRows; r++) {
      let codeCol = null;
      let nameCol = null;
      const scanCols = Math.min(maxCol, 8);
      for (let c = 1; c <= scanCols; c++) {
        const s = norm(valAt(values, r, c)).replace(/ /g, '');
        if (!s) continue;
        if (codeCol === null && HEADER_KEYS.some(function (k) { return s.indexOf(k) >= 0; })) codeCol = c;
        if (nameCol === null && NAME_KEYS.some(function (k) { return s.indexOf(k) >= 0; })) nameCol = c;
      }
      if (codeCol && nameCol) return { headerRow: r, codeCol: codeCol, nameCol: nameCol };
      if (codeCol && best === null) best = { headerRow: r, codeCol: codeCol, nameCol: codeCol + 1 };
    }
    if (best) return best;
    // [v2] 无「序号/名称」表头时的兜底：纯数字列通常是数量或金额，不能当编号列；
    // 且至少要有 3 个文本编号才认为找到了编号体系，否则标记为低置信度（不分级）。
    const counts = {};
    const maxC = Math.min(maxCol, 4);
    const maxR = Math.min(maxRow, 200);
    for (let c = 1; c <= maxC; c++) {
      for (let r = 1; r <= maxR; r++) {
        const fam = codeFamily(valAt(values, r, c)).family;
        if (fam && fam !== 'part' && fam !== 'num') counts[c] = (counts[c] || 0) + 1;
      }
    }
    let col = 1;
    let max = 0;
    Object.keys(counts).forEach(function (k) {
      if (counts[k] > max) { max = counts[k]; col = parseInt(k, 10); }
    });
    if (max < 3) return { headerRow: 3, codeCol: 1, nameCol: 2, lowConfidence: true };
    return { headerRow: 3, codeCol: col, nameCol: col + 1 };
  }

  function isBannerRow(values, r, nameCol, maxCol) {
    for (let i = 0; i < 3; i++) {
      const c = nameCol + i;
      if (c < 1 || c > maxCol) continue;
      if (codeFamily(valAt(values, r, c)).family === 'part') return true;
    }
    return false;
  }

  // ---------------------------------------------------------------- 层级推导

  function detectLevels(values, maxRow, maxCol) {
    const header = findHeader(values, maxRow, maxCol);
    const headerRow = header.headerRow;
    const codeCol = header.codeCol;
    const nameCol = header.nameCol;
    // [v2] 指标表等没有编号体系的表不做分级，避免把数值列误当编号列后强行分组
    if (header.lowConfidence) {
      return {
        headerRow: headerRow, codeCol: codeCol, nameCol: nameCol,
        levels: {}, maxLevel: 0, partIsFamily: false, families: [], familyList: [],
      };
    }
    let start = headerRow + 1;

    let partCount = 0;
    const seen = new Set();
    for (let r = start; r <= maxRow; r++) {
      const fam = codeFamily(valAt(values, r, codeCol)).family;
      if (fam === 'part') partCount++;
      if (fam) seen.add(fam);
    }
    const partIsFamily = partCount >= 2;

    const levels = {};
    let firstStruct = null;
    for (let r = start; r <= maxRow; r++) {
      if (isBannerRow(values, r, nameCol, maxCol) || codeFamily(valAt(values, r, codeCol)).family) {
        firstStruct = r;
        break;
      }
    }
    if (firstStruct) {
      for (let r = start; r < firstStruct; r++) levels[r] = 0;
      start = firstStruct;
    }

    const famLevel = {};
    const famRow = {};
    const famByOrd = {};
    // [v2] 段落起点：遇到段级编号重置，点号编号只在本段内找父级，避免跨段串项
    let sectionStart = start;
    let lastCodeLevel = 0;
    let lastLevel = 0;

    function remember(family, r, lv, info) {
      famLevel[family] = lv;
      famRow[family] = r;
      const ord = codeOrdinal(family, info && info.text);
      if (ord !== null && ord !== undefined) {
        if (!famByOrd[family]) famByOrd[family] = {};
        famByOrd[family][ord] = { row: r, level: lv };
      }
    }

    for (let r = start; r <= maxRow; r++) {
      const info = codeFamily(valAt(values, r, codeCol));
      const family = info.family;
      const nameVal = norm(valAt(values, r, nameCol));

      if (isBannerRow(values, r, nameCol, maxCol) || (family === 'part' && !partIsFamily)) {
        levels[r] = 0;
        lastCodeLevel = 0;
        lastLevel = 0;
        if (family === 'part') sectionStart = r;
        continue;
      }

      if (family) {
        // [v2] 段级编号出现即开启新段，后续点号编号只与本段内的父级配对
        if (family === 'roman' || family === 'part' || family === 'cn' || family === 'cn_paren') {
          sectionStart = r;
        }
        let lv;
        if (family === 'roman' || family === 'part') {
          lv = 0;
        } else if (family === 'dotted') {
          // 点号编号（3.1 / 3.1.2）按“首段数字”找上级编号定位，
          // 不能再以上一行层级为基准，否则同级条目会逐行越陷越深。
          const parts = info.parts || [];
          const first = parts.length ? parts[0] : null;
          const dotDepth = Math.max(1, parts.length - 1);
          const parentFams = ['num', 'cn_paren', 'cn', 'roman', 'part', 'num_paren', 'num_ju', 'circle'];
          let base = null;
          let baseRow = -1;
          if (first !== null && !isNaN(first)) {
            for (let i = 0; i < parentFams.length; i++) {
              const rec = famByOrd[parentFams[i]] && famByOrd[parentFams[i]][first];
              if (rec && rec.row < r && rec.row >= sectionStart && rec.row > baseRow) {
                baseRow = rec.row;
                base = rec.level;
              }
            }
          }
          if (base === null) {
            for (let i = 0; i < parentFams.length; i++) {
              const f = parentFams[i];
              if (famRow[f] !== undefined && famRow[f] < r && famRow[f] >= sectionStart && famRow[f] > baseRow) {
                baseRow = famRow[f];
                base = famLevel[f];
              }
            }
          }
          // [v2] 段内确实没有可用父级时，退回全局最近的段级编号，避免沿用上一段的层级
          if (base === null) {
            for (let i = 0; i < parentFams.length; i++) {
              const f = parentFams[i];
              if (famRow[f] !== undefined && famRow[f] < r && famRow[f] > baseRow) {
                baseRow = famRow[f];
                base = famLevel[f];
              }
            }
          }
          lv = Math.min(MAX_LEVEL, (base === null ? 0 : base) + dotDepth);
        } else {
          const parents = FAMILY_PARENTS[family] || [];
          const known = parents
            .filter(function (p) { return famLevel[p] !== undefined; })
            .map(function (p) { return [famRow[p] === undefined ? -1 : famRow[p], famLevel[p]]; });
          known.sort(function (a, b) { return a[0] - b[0]; });
          lv = known.length ? Math.min(MAX_LEVEL, known[known.length - 1][1] + 1) : 0;
        }
        levels[r] = lv;
        remember(family, r, lv, info);
        lastCodeLevel = lv;
        lastLevel = lv;
        continue;
      }

      if (rowIsEmpty(values, r, maxCol)) {
        let laterHasData = false;
        const stop = Math.min(maxRow, r + 30);
        for (let rr = r + 1; rr <= stop; rr++) {
          if (!rowIsEmpty(values, rr, maxCol)) { laterHasData = true; break; }
        }
        levels[r] = laterHasData ? lastLevel : 0;
        continue;
      }

      let lv;
      if (SECTION_TOTAL_WORDS.some(function (w) { return nameVal.indexOf(w) >= 0; })) {
        lv = famLevel.cn !== undefined ? famLevel.cn : 0;
      } else {
        lv = Math.min(MAX_LEVEL, lastCodeLevel + 1);
      }
      levels[r] = lv;
      lastLevel = lv;
    }

    let maxLevel = 0;
    Object.keys(levels).forEach(function (k) {
      if (levels[k] > maxLevel) maxLevel = levels[k];
    });

    return {
      headerRow: headerRow,
      codeCol: codeCol,
      nameCol: nameCol,
      levels: levels,
      maxLevel: maxLevel,
      partIsFamily: partIsFamily,
      families: Array.from(seen),
      familyList: CODE_RULES.map(function (r) { return r[0]; }).filter(function (f) { return seen.has(f); }),
    };
  }

  function detectSheetValues(values, maxRow, maxCol) {
    const last = lastDataRow(values, maxRow, maxCol);
    const det = detectLevels(values, last || maxRow, maxCol);
    det.lastDataRow = last || maxRow;
    let structured = 0;
    Object.keys(det.levels).forEach(function (k) { if (det.levels[k] > 0) structured++; });
    det.structuredRows = structured;
    return det;
  }

  // ---------------------------------------------------------------- 层级求和校验

  const SUM_AMOUNT_KEYS = ['合计', '金额', '费用', '投资', '造价', '含税', '税', '设备费', '安装费', '建筑工程费', '安装工程费', '设备购置费', '其他费用', '预备费', '静态投资', '总投资', '价差', '利息', '资金', '成本', '收入', '支出', '直接费', '间接费', '利润', '税金', '建筑及安装', '补偿'];
  const SUM_EXCLUDE_KEYS = ['单价', '数量', '单位', '费率', '比例', '税率', '定额', '指标', '系数', '备注', '序号', '编号', '名称', '日期', '年月', '%', '％'];
  const SUM_SUBTOTAL_RE = /小计|合计|总计|累计|总额|总概算|静态投资|总投资|建设投资|工程造价|合价|^一至/;
  const SUM_FEE_RE = /增值税|基本预备费|价差预备费|建设期利息|预备费|^税金$/;
  const SUM_SUMMARY_RE = /分年度|完成工作量|年度完成|完成投资/;
  const SUM_BANNER_RE = /^第[一二三四五六七八九十百零〇\d]{1,4}\s*(项|部分|章|节|篇)/;
  const MAX_SUM_ERRORS = 500;

  function toNumberValue(value) {
    if (typeof value === 'number') {
      return isFinite(value) ? { ok: true, value: value } : { ok: false };
    }
    if (typeof value !== 'string') return { ok: false };
    let s = value.trim();
    if (!s) return { ok: false };
    if (/^#[A-Za-z0-9\/_!]+$/.test(s)) return { ok: false, error: s };
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    s = s.replace(/[,，\s\u00a0\u3000]/g, '').replace(/^[￥¥]/, '');
    if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(s)) return { ok: false };
    const n = parseFloat(s);
    if (!isFinite(n)) return { ok: false };
    return { ok: true, value: neg ? -n : n };
  }

  function cellErrorValue(value) {
    if (typeof value !== 'string') return null;
    const s = value.trim();
    if (!/^#[A-Za-z0-9\/_!]+$/.test(s)) return null;
    return /^#(REF|VALUE|DIV\/0|NAME|N\/A|NUM|NULL|SPILL|CALC|GETTING_DATA)!?$/i.test(s) ? s : null;
  }

  function isAmountHeaderLabel(label) {
    const s = norm(label).replace(/\s+/g, '');
    if (!s) return false;
    for (let i = 0; i < SUM_EXCLUDE_KEYS.length; i++) {
      if (s.indexOf(SUM_EXCLUDE_KEYS[i]) >= 0) return false;
    }
    for (let i = 0; i < SUM_AMOUNT_KEYS.length; i++) {
      if (s.indexOf(SUM_AMOUNT_KEYS[i]) >= 0) return true;
    }
    return /第?\-?\d+年/.test(s);
  }

  function isSumBanner(value) {
    const s = norm(value);
    return !!s && s.length <= 40 && SUM_BANNER_RE.test(s);
  }

  function sumHeaderRows(values, headerRow, last, codeCol, nameCol, maxCol) {
    const rows = [headerRow];
    for (let r = headerRow + 1; r <= Math.min(headerRow + 2, last); r++) {
      if (norm(valAt(values, r, codeCol))) break;
      let banner = false;
      for (let i = 0; i < 3; i++) {
        if (isSumBanner(valAt(values, r, nameCol + i))) { banner = true; break; }
      }
      if (banner) break;
      let textCount = 0;
      for (let c = 1; c <= maxCol; c++) {
        const v = valAt(values, r, c);
        if (typeof v === 'string' && norm(v) && !toNumberValue(v).ok && !cellErrorValue(v)) textCount++;
      }
      if (textCount < 2) break;
      rows.push(r);
    }
    return rows;
  }

  function sumHeaderText(values, rows, c) {
    const parts = [];
    for (let i = 0; i < rows.length; i++) {
      const v = norm(valAt(values, rows[i], c));
      if (v) parts.push(v);
    }
    return parts.join(' ');
  }

  function validateSheetSums(sheet) {
    const empty = {
      status: 'na', reason: '没有可校验的数据', levelSource: 'none', maxLevel: 0,
      checkedCells: 0, checkedRows: 0, errorCount: 0, errors: [], errorGroups: [],
      errorRows: 0, maxDiff: 0, columns: [], truncated: false, hiddenErrors: 0,
    };
    if (!sheet) return empty;

    const values = sheet.values || {};
    const detection = sheet.detection || {};
    const outline = sheet.outline || { rows: {}, max: 0 };
    const lastDataRow = detection.lastDataRow || sheet.maxRow || 0;
    const maxCol = sheet.maxCol || 0;
    const headerRow = detection.headerRow;
    const codeCol = detection.codeCol;
    const nameCol = detection.nameCol;

    if (!lastDataRow || !maxCol || !headerRow || !codeCol || !nameCol) {
      return Object.assign({}, empty, { reason: '未识别到表头或数据行' });
    }
    if ((!detection.maxLevel || detection.maxLevel <= 0) && (!detection.structuredRows || detection.structuredRows <= 0)) {
      return Object.assign({}, empty, {
        reason: '未识别到编号层级，无法判断父子关系。原表分级可能不是按编号生成的。',
      });
    }

    const headerRows = sumHeaderRows(values, headerRow, lastDataRow, codeCol, nameCol, maxCol);
    const headerEnd = headerRows[headerRows.length - 1];
    const headerLabels = {};
    for (let c = 1; c <= maxCol; c++) headerLabels[c] = sumHeaderText(values, headerRows, c);

    const amountCols = [];
    for (let c = 1; c <= maxCol; c++) {
      if (isAmountHeaderLabel(headerLabels[c])) amountCols.push(c);
    }
    if (!amountCols.length) {
      return Object.assign({}, empty, { reason: '未找到金额列（数量、单价、费率等列已跳过）' });
    }

    const outlineLevels = {};
    if (outline.max > 0) {
      const orows = outline.rows || {};
      for (let r = headerEnd + 1; r <= lastDataRow; r++) {
        outlineLevels[r] = (orows[r] && orows[r].level) || 0;
      }
    }

    function compute(levels, source) {
      const data = [];
      for (let r = headerEnd + 1; r <= lastDataRow; r++) {
        const row = values[r];
        if (!row) continue;
        let code = norm(row[codeCol]);
        let name = norm(row[nameCol]);
        if (!name) {
          for (let i = 1; i <= 2; i++) {
            const v = norm(row[nameCol + i]);
            if (v) { name = v; break; }
          }
        }
        const nums = {};
        const errs = {};
        let numeric = 0;
        for (let ci = 0; ci < amountCols.length; ci++) {
          const c = amountCols[ci];
          const raw = row[c];
          if (raw === undefined || raw === null || raw === '') continue;
          const err = cellErrorValue(raw);
          if (err) { errs[c] = err; continue; }
          const num = toNumberValue(raw);
          if (num.ok) { nums[c] = num.value; numeric++; }
        }
        if (!code && !name && !numeric && !Object.keys(errs).length) continue;
        const label = name || code;
        data.push({
          r: r,
          lv: levels[r] || 0,
          code: code,
          name: label,
          banner: isSumBanner(code) || isSumBanner(name),
          sub: SUM_SUBTOTAL_RE.test(label),
          fee: SUM_FEE_RE.test(label),
          summary: SUM_SUMMARY_RE.test(label),
          nums: nums,
          errs: errs,
        });
      }

      const n = data.length;
      const children = {};
      const stack = [];
      let bannerIdx = -1;
      for (let i = 0; i < n; i++) {
        const lv = data[i].lv;
        if (data[i].banner && lv === 0) {
          bannerIdx = i;
          stack.length = 0;
          stack.push(i);
          continue;
        }
        if (lv === 0) {
          if (bannerIdx >= 0) {
            if (!children[bannerIdx]) children[bannerIdx] = [];
            children[bannerIdx].push(i);
          }
          stack.length = 0;
          stack.push(i);
          continue;
        }
        while (stack.length && data[stack[stack.length - 1]].lv >= lv) stack.pop();
        const p = stack.length ? stack[stack.length - 1] : -1;
        if (p >= 0) {
          if (!children[p]) children[p] = [];
          children[p].push(i);
        }
        stack.push(i);
      }

      const errors = [];
      let hiddenErrors = 0;
      let checkedCells = 0;
      let checkedRows = 0;
      let maxDiff = 0;
      let maxLevel = 0;

      function pushError(e) {
        if (errors.length < MAX_SUM_ERRORS) errors.push(e);
        else hiddenErrors++;
      }

      for (let i = 0; i < n; i++) {
        if (data[i].lv > maxLevel) maxLevel = data[i].lv;
      }

      for (const pKey in children) {
        const p = parseInt(pKey, 10);
        if (!children[p] || !children[p].length) continue;
        if (data[p].sub) continue;

        const kids = children[p];
        const allKids = [];
        for (let i = 0; i < kids.length; i++) {
          if (!data[kids[i]].sub) allKids.push(kids[i]);
        }
        if (!allKids.length) continue;

        const coded = [];
        const feeExtra = [];
        const uncodedSummary = {};
        for (let i = 0; i < allKids.length; i++) {
          const k = allKids[i];
          const hasCode = !!codeFamily(data[k].code).family;
          if (hasCode) coded.push(k);
          if (data[k].fee && !hasCode) feeExtra.push(k);
          if (data[k].summary && !hasCode) uncodedSummary[k] = true;
        }

        const candidates = [];
        const seen = {};
        function addCandidate(list, label) {
          const uniq = [];
          const used = {};
          for (let i = 0; i < list.length; i++) {
            if (!used[list[i]]) { used[list[i]] = true; uniq.push(list[i]); }
          }
          uniq.sort(function (a, b) { return a - b; });
          const key = uniq.join(',');
          if (seen[key]) return;
          seen[key] = true;
          candidates.push({ list: uniq, label: label });
        }

        if (coded.length) {
          if (feeExtra.length) addCandidate(coded.concat(feeExtra), '编号子项 + 费用行');
          addCandidate(coded, '编号子项');
          const contribAll = [];
          for (let i = 0; i < allKids.length; i++) {
            if (!uncodedSummary[allKids[i]]) contribAll.push(allKids[i]);
          }
          addCandidate(contribAll, '全部子项');
        } else {
          addCandidate(allKids, '全部子项');
        }

        let rowChecked = false;
        for (let ci = 0; ci < amountCols.length; ci++) {
          const c = amountCols[ci];
          const pv = data[p].nums[c];
          if (pv === undefined) continue;
          const opts = [];
          for (let oi = 0; oi < candidates.length; oi++) {
            const list = candidates[oi].list;
            let sum = 0;
            let count = 0;
            for (let i = 0; i < list.length; i++) {
              const v = data[list[i]].nums[c];
              if (v !== undefined) { sum += v; count++; }
            }
            if (count >= 2) opts.push({ sum: sum, count: count, label: candidates[oi].label, list: list });
          }
          if (!opts.length) continue;
          let best = opts[0];
          for (let oi = 1; oi < opts.length; oi++) {
            if (Math.abs(pv - opts[oi].sum) < Math.abs(pv - best.sum)) best = opts[oi];
          }
          checkedCells++;
          rowChecked = true;
          const diff = pv - best.sum;
          const tol = Math.max(0.05, Math.abs(pv) * 1e-9);
          if (Math.abs(diff) > tol) {
            if (Math.abs(diff) > maxDiff) maxDiff = Math.abs(diff);
            const fRow = sheet.formulas && sheet.formulas[data[p].r];
            const hasFormula = !!(fRow && fRow[c] !== undefined);
            let hint = '';
            const matchTol = Math.max(0.05, Math.abs(diff) * 1e-9);
            for (let li = 0; li < best.list.length; li++) {
              const k = best.list[li];
              const kv = data[k].nums[c];
              if (kv === undefined) continue;
              if (Math.abs(Math.abs(diff) - Math.abs(kv)) <= matchTol) {
                hint = '差额 = 子项「' + (data[k].name || '（无名称）') + '」，疑似重复计入或漏计';
                break;
              }
            }
            pushError({
              kind: 'sum',
              row: data[p].r,
              col: c,
              letter: colLetter(c),
              rowName: data[p].name || '（无名称）',
              colLabel: headerLabels[c] || colLetter(c),
              parentValue: pv,
              childSum: best.sum,
              diff: diff,
              childCount: best.count,
              basis: best.label,
              hint: hint,
              level: data[p].lv,
              hasFormula: hasFormula,
              formula: hasFormula ? String(fRow[c] || '') : '',
            });
          }
        }
        if (rowChecked) checkedRows++;
      }

      for (let i = 0; i < data.length; i++) {
        for (let ci = 0; ci < amountCols.length; ci++) {
          const c = amountCols[ci];
          const err = data[i].errs[c];
          if (!err) continue;
          const fRow = sheet.formulas && sheet.formulas[data[i].r];
          const hasFormula = !!(fRow && fRow[c] !== undefined);
          pushError({
            kind: 'cell',
            row: data[i].r,
            col: c,
            letter: colLetter(c),
            rowName: data[i].name || '（无名称）',
            colLabel: headerLabels[c] || colLetter(c),
            parentValue: null,
            childSum: null,
            diff: 0,
            childCount: 0,
            basis: '',
            level: data[i].lv,
            hasFormula: hasFormula,
            formula: hasFormula ? String(fRow[c] || '') : '',
            message: err,
          });
        }
      }

      errors.sort(function (a, b) {
        if (a.row !== b.row) return a.row - b.row;
        return a.col - b.col;
      });

      const groupMap = {};
      for (let i = 0; i < errors.length; i++) {
        const e = errors[i];
        let g = groupMap[e.row];
        if (!g) {
          g = groupMap[e.row] = { row: e.row, rowName: e.rowName, level: e.level, cells: [] };
        }
        g.cells.push(e);
      }
      const errorGroups = Object.keys(groupMap).map(function (k) {
        return groupMap[k];
      }).sort(function (a, b) { return a.row - b.row; });

      const columns = amountCols.map(function (c) {
        return { col: c, letter: colLetter(c), label: headerLabels[c] || colLetter(c) };
      });

      let status = 'pass';
      let reason = '';
      if (errors.length || hiddenErrors) {
        status = 'diff';
      } else if (!checkedCells) {
        status = 'na';
        reason = '没有找到“父行 + 至少两个子项”的可校验结构';
      }

      return {
        status: status,
        reason: reason,
        levelSource: source,
        maxLevel: maxLevel,
        checkedCells: checkedCells,
        checkedRows: checkedRows,
        errorCount: errors.length + hiddenErrors,
        errors: errors,
        errorGroups: errorGroups,
        errorRows: errorGroups.length,
        maxDiff: maxDiff,
        columns: columns,
        truncated: hiddenErrors > 0,
        hiddenErrors: hiddenErrors,
      };
    }

    let result;
    if (outline.max > 0) {
      result = compute(outlineLevels, 'outline');
      if (!result.checkedCells && detection.maxLevel > 0 && detection.levels) {
        const detected = compute(detection.levels, 'detected');
        if (detected.checkedCells) {
          detected.fallback = true;
          result = detected;
        }
      }
    } else {
      result = compute(detection.levels || {}, 'detected');
    }
    return result;
  }

  // ---------------------------------------------------------------- 读取分级

  function readOutline(xmlText) {
    const rows = {};
    let max = 0;
    let summaryBelow = true;
    const re = /<row\b([^>]*?)(\/?)>/g;
    let m;
    while ((m = re.exec(xmlText))) {
      const attrs = m[1];
      const r = parseInt(attrOf(attrs, 'r') || '0', 10);
      if (!r) continue;
      const level = parseInt(attrOf(attrs, 'outlineLevel') || '0', 10) || 0;
      const hidden = /^(1|true)$/.test(attrOf(attrs, 'hidden') || '');
      const collapsed = /^(1|true)$/.test(attrOf(attrs, 'collapsed') || '');
      if (level || hidden || collapsed) rows[r] = { level: level, hidden: hidden, collapsed: collapsed };
      if (level > max) max = level;
    }
    const op = /<outlinePr\b[^>]*>/.exec(xmlText);
    if (op) summaryBelow = !/summaryBelow="(0|false)"/.test(op[0]);
    return { rows: rows, max: max, summaryBelow: summaryBelow };
  }

  function outlineDist(outline) {
    const dist = {};
    Object.keys(outline.rows).forEach(function (r) {
      const lv = outline.rows[r].level;
      if (lv > 0) dist[lv] = (dist[lv] || 0) + 1;
    });
    return dist;
  }

  // ---------------------------------------------------------------- 写入分级

  function applyOutline(xmlText, rowMap, maxLevel, options) {
    options = options || {};
    const clearColumns = options.clearColumns !== false;
    const summaryBelow = options.summaryBelow === true;
    const sb = summaryBelow ? '1' : '0';

    let xml = xmlText;

    xml = xml.replace(/<row\b([^>]*?)(\/?)>/g, function (full, attrs, selfClose) {
      const r = parseInt(attrOf(attrs, 'r') || '0', 10);
      let a = attrs.replace(/\s+(?:outlineLevel|hidden|collapsed)="[^"]*"/g, '').replace(/\s+$/, '');
      if (r) {
        const info = rowMap[r];
        if (info) {
          if (info.level > 0) a += ' outlineLevel="' + info.level + '"';
          if (info.hidden) a += ' hidden="1"';
          if (info.collapsed) a += ' collapsed="1"';
        }
      }
      return '<row' + a + selfClose + '>';
    });

    if (clearColumns) {
      xml = xml.replace(/<col\b([^>]*?)(\/?)>/g, function (full, attrs, selfClose) {
        const a = attrs.replace(/\s+(?:outlineLevel|hidden|collapsed)="[^"]*"/g, '').replace(/\s+$/, '');
        return '<col' + a + selfClose + '>';
      });
    }

    const lv = Math.max(0, Math.min(MAX_LEVEL, maxLevel || 0));
    if (/<sheetFormatPr\b/.test(xml)) {
      xml = xml.replace(/<sheetFormatPr\b([^>]*?)(\/?)>/, function (full, attrs, selfClose) {
        let a = attrs.replace(/\s+outlineLevelRow="[^"]*"/g, '').replace(/\s+outlineLevelCol="[^"]*"/g, '').replace(/\s+$/, '');
        a += ' outlineLevelRow="' + lv + '"';
        return '<sheetFormatPr' + a + selfClose + '>';
      });
    } else {
      xml = xml.replace(/(<sheetData\b)/, '<sheetFormatPr outlineLevelRow="' + lv + '"/>$1');
    }

    if (/<outlinePr\b/.test(xml)) {
      xml = xml.replace(/<outlinePr\b([^>]*?)(\/?)>/, function (full, attrs, selfClose) {
        let a = attrs.replace(/\s+summaryBelow="[^"]*"/g, '').replace(/\s+$/, '');
        a += ' summaryBelow="' + sb + '"';
        return '<outlinePr' + a + selfClose + '>';
      });
    } else if (/<sheetPr\b[^>]*\/>/.test(xml)) {
      xml = xml.replace(/<sheetPr\b([^>]*?)\/>/, function (full, attrs) {
        return '<sheetPr' + attrs + '><outlinePr summaryBelow="' + sb + '"/></sheetPr>';
      });
    } else if (/<sheetPr\b[^>]*>/.test(xml)) {
      if (/<pageSetUpPr\b/.test(xml)) {
        xml = xml.replace(/<pageSetUpPr\b/, '<outlinePr summaryBelow="' + sb + '"/><pageSetUpPr');
      } else {
        xml = xml.replace(/<\/sheetPr>/, '<outlinePr summaryBelow="' + sb + '"/></sheetPr>');
      }
    } else {
      xml = xml.replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr><outlinePr summaryBelow="' + sb + '"/></sheetPr>');
    }

    return xml;
  }

  function makeRowMap(levels) {
    const map = {};
    Object.keys(levels || {}).forEach(function (k) {
      const r = parseInt(k, 10);
      const lv = Math.max(0, Math.min(MAX_LEVEL, parseInt(levels[k], 10) || 0));
      if (r && lv > 0) map[r] = { level: lv, hidden: false, collapsed: false };
    });
    return map;
  }

  function maxOf(levels) {
    let m = 0;
    Object.keys(levels || {}).forEach(function (k) {
      if (levels[k] > m) m = levels[k];
    });
    return Math.min(MAX_LEVEL, m);
  }

  // 全量重建：applyOutline 会先剥掉每一行旧的 outlineLevel/hidden/collapsed，
  // 再按 rowMap 写入，并清列分级、重设 sheetFormatPr / outlinePr。
  // 因此本函数天然包含「取消原有分级」，无需在调用前额外执行 applyClearOutline。
  function applyAutoOutline(xmlText, levels) {
    const rowMap = makeRowMap(levels);
    const maxLevel = maxOf(levels);
    return {
      xml: applyOutline(xmlText, rowMap, maxLevel, { summaryBelow: false }),
      maxLevel: maxLevel,
    };
  }

  function applyClearOutline(xmlText, summaryBelow) {
    return applyOutline(xmlText, {}, 0, { summaryBelow: summaryBelow === true });
  }

  // ---------------------------------------------------------------- 关系解析

  function parseRelationships(xml) {
    const out = {};
    const re = /<Relationship\b[^>]*>/g;
    let m;
    while ((m = re.exec(xml))) {
      const tag = m[0];
      const id = attrOf(tag, 'Id');
      if (!id) continue;
      out[id] = {
        id: id,
        target: attrOf(tag, 'Target'),
        type: attrOf(tag, 'Type'),
        targetMode: attrOf(tag, 'TargetMode'),
      };
    }
    return out;
  }

  function normalizePath(path) {
    const parts = path.split('/');
    const out = [];
    parts.forEach(function (p) {
      if (!p || p === '.') return;
      if (p === '..') out.pop();
      else out.push(p);
    });
    return out.join('/');
  }

  function resolvePart(target, basePath) {
    if (!target) return null;
    if (target.charAt(0) === '/') return normalizePath(target.slice(1));
    const baseDir = basePath ? basePath.split('/').slice(0, -1).join('/') : '';
    return normalizePath((baseDir ? baseDir + '/' : '') + target);
  }

  // ---------------------------------------------------------------- 工作簿

  async function getWorkbookMeta(zip) {
    const wbPath = 'xl/workbook.xml';
    const relsPath = 'xl/_rels/workbook.xml.rels';
    const wbFile = zip.file(wbPath);
    const relsFile = zip.file(relsPath);
    if (!wbFile || !relsFile) throw new Error('缺少 xl/workbook.xml，可能不是有效的 xlsx 文件');
    const wbXml = await wbFile.async('string');
    const relsXml = await relsFile.async('string');
    const rels = parseRelationships(relsXml);
    const sheets = [];
    let idx = 0;
    const re = /<sheet\s[^>]*>/g;
    let m;
    while ((m = re.exec(wbXml))) {
      const tag = m[0];
      const name = attrOf(tag, 'name');
      const rid = attrOf(tag, 'r:id') || attrOf(tag, 'relationships:id');
      if (!name || !rid || !rels[rid]) { idx++; continue; }
      const path = resolvePart(rels[rid].target, wbPath);
      sheets.push({
        index: idx,
        name: name,
        sheetId: parseInt(attrOf(tag, 'sheetId') || '0', 10),
        state: attrOf(tag, 'state') || 'visible',
        rid: rid,
        path: path,
        isWorksheet: /^xl\/worksheets\/[^/]+\.xml$/.test(path),
      });
      idx++;
    }
    return { wbPath: wbPath, relsPath: relsPath, wbXml: wbXml, relsXml: relsXml, rels: rels, sheets: sheets };
  }

  async function loadWorkbook(zip, onProgress) {
    const meta = await getWorkbookMeta(zip);
    if (onProgress) onProgress('正在读取共享字符串…');
    const sharedStrings = await readSharedStrings(zip);
    const sheets = [];
    for (let i = 0; i < meta.sheets.length; i++) {
      const info = meta.sheets[i];
      if (onProgress) onProgress('正在识别「' + info.name + '」…');
      if (!info.isWorksheet) {
        sheets.push({
          name: info.name, path: info.path, rid: info.rid, sheetId: info.sheetId,
          state: info.state, isWorksheet: false, xml: '', originalXml: '',
          outline: { rows: {}, max: 0, summaryBelow: true },
          originalOutline: { rows: {}, max: 0, summaryBelow: true },
          detection: { levels: {}, maxLevel: 0, structuredRows: 0, lastDataRow: 0, headerRow: 0, codeCol: 1, nameCol: 2, familyList: [], families: [] },
          values: {}, formulas: {}, maxRow: 0, maxCol: 0, modified: false, isCopy: false,
          validation: { status: 'na', reason: '非数据表', levelSource: 'none', maxLevel: 0, checkedCells: 0, checkedRows: 0, errorCount: 0, errors: [], errorGroups: [], errorRows: 0, maxDiff: 0, columns: [], truncated: false, hiddenErrors: 0 },
        });
        continue;
      }
      const xml = await zip.file(info.path).async('string');
      const parsed = readSheetValues(xml, sharedStrings);
      const detection = detectSheetValues(parsed.rows, parsed.maxRow, parsed.maxCol);
      const outline = readOutline(xml);
      const sheet = {
        name: info.name, path: info.path, rid: info.rid, sheetId: info.sheetId,
        state: info.state, isWorksheet: true, xml: xml, originalXml: xml,
        outline: outline, originalOutline: readOutline(xml),
        detection: detection, values: parsed.rows, formulas: parsed.formulas || {},
        maxRow: detection.lastDataRow, maxCol: parsed.maxCol,
        modified: false, isCopy: false,
        validation: null,
      };
      sheet.validation = validateSheetSums(sheet);
      sheets.push(sheet);
    }
    return { meta: meta, sharedStrings: sharedStrings, sheets: sheets };
  }

  // ---------------------------------------------------------------- 复制工作表

  function relsPathFor(sheetPath) {
    return sheetPath.replace(/([^/]+)$/, '_rels/$1.rels');
  }

  function validateSheetName(name, existing) {
    if (!name || !name.trim()) throw new Error('工作表名不能为空');
    if (name.length > 31) throw new Error('工作表名最多 31 个字符');
    if (/[\[\]:*?\/\\]/.test(name)) throw new Error('工作表名不能包含 [ ] : * ? / \\');
    if (/^'|'$/.test(name)) throw new Error('工作表名不能以单引号开头或结尾');
    if (existing.some(function (n) { return n.toLowerCase() === name.toLowerCase(); })) {
      throw new Error('工作表名已存在：' + name);
    }
  }

  function makeCopyName(base, existing) {
    let n = 1;
    while (n < 1000) {
      const suffix = n === 1 ? '（副本）' : '（副本' + n + '）';
      let root = base;
      if (root.length + suffix.length > 31) root = root.slice(0, 31 - suffix.length);
      const name = root + suffix;
      if (!existing.some(function (x) { return x.toLowerCase() === name.toLowerCase(); })) return name;
      n++;
    }
    throw new Error('无法生成副本名称');
  }

  function uniqueSheetPath(zip) {
    let max = 0;
    zip.forEach(function (path) {
      const m = /^xl\/worksheets\/sheet(\d+)\.xml$/.exec(path);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
    let n = max + 1;
    while (zip.file('xl/worksheets/sheet' + n + '.xml')) n++;
    return 'xl/worksheets/sheet' + n + '.xml';
  }

  async function clonePrinterSettings(zip, srcPath, rel) {
    const abs = resolvePart(rel.target, srcPath);
    const source = zip.file(abs);
    if (!source) return rel.target;
    let max = 0;
    zip.forEach(function (p) {
      const m = /^xl\/printerSettings\/printerSettings(\d+)\.bin$/.exec(p);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
    const newAbs = 'xl/printerSettings/printerSettings' + (max + 1) + '.bin';
    zip.file(newAbs, await source.async('uint8array'));
    addContentTypeOverride(zip, newAbs, 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings');
    return '../printerSettings/' + newAbs.split('/').pop();
  }

  function replaceRelTarget(xml, id, newTarget) {
    return xml.replace(/<Relationship\b[^>]*>/g, function (tag) {
      if (attrOf(tag, 'Id') !== id) return tag;
      return tag.replace(/(\sTarget=")[^"]*(")/, '$1' + escapeXml(newTarget) + '$2');
    });
  }

  async function addContentTypeOverride(zip, partPath, contentType) {
    const path = '[Content_Types].xml';
    let ct = await zip.file(path).async('string');
    if (ct.indexOf('PartName="/' + partPath + '"') >= 0) return;
    if (/Default\s+Extension="bin"/i.test(ct) && /printerSettings/.test(contentType)) return;
    ct = ct.replace('</Types>', '<Override PartName="/' + partPath + '" ContentType="' + contentType + '"/></Types>');
    zip.file(path, ct);
  }

  async function duplicateSheet(zip, meta, srcName) {
    const src = meta.sheets.find(function (s) { return s.name === srcName; });
    if (!src) throw new Error('找不到工作表：' + srcName);
    if (!src.isWorksheet) throw new Error('该工作表不是普通数据表，无法复制');
    const newName = makeCopyName(src.name, meta.sheets.map(function (s) { return s.name; }));
    validateSheetName(newName, meta.sheets.map(function (s) { return s.name; }));

    const srcXml = await zip.file(src.path).async('string');
    const newPath = uniqueSheetPath(zip);
    let newXml = srcXml
      .replace(/(<sheetPr\b[^>]*?)\s+codeName="[^"]*"/, '$1')
      .replace(/\s+tabSelected="1"/g, '');

    // 复制工作表关系：只允许“打印设置 / 超链接”两类，其它嵌入式对象直接拒绝，避免生成损坏文件
    const srcRelsPath = relsPathFor(src.path);
    let newRelsXml = null;
    if (zip.file(srcRelsPath)) {
      let relsXml = await zip.file(srcRelsPath).async('string');
      const entries = parseRelationships(relsXml);
      const ids = Object.keys(entries);
      for (let i = 0; i < ids.length; i++) {
        const rel = entries[ids[i]];
        if (!/\/printerSettings$/.test(rel.type) && !/\/hyperlink$/.test(rel.type)) {
          throw new Error('「' + src.name + '」含有图片/图表/批注/表格对象，直接复制可能导致文件损坏。请先在 Excel 中手动复制该表，或仅做分级操作。');
        }
      }
      for (let i = 0; i < ids.length; i++) {
        const rel = entries[ids[i]];
        if (/\/printerSettings$/.test(rel.type)) {
          const newTarget = await clonePrinterSettings(zip, src.path, rel);
          relsXml = replaceRelTarget(relsXml, rel.id, newTarget);
        }
      }
      newRelsXml = relsXml;
    }

    zip.file(newPath, newXml);
    if (newRelsXml !== null) zip.file(relsPathFor(newPath), newRelsXml);

    const ctPath = '[Content_Types].xml';
    let ct = await zip.file(ctPath).async('string');
    if (ct.indexOf('PartName="/' + newPath + '"') < 0) {
      ct = ct.replace('</Types>', '<Override PartName="/' + newPath + '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
      zip.file(ctPath, ct);
    }

    let maxRid = 0;
    Object.keys(meta.rels).forEach(function (id) {
      const m = /^rId(\d+)$/.exec(id);
      if (m) maxRid = Math.max(maxRid, parseInt(m[1], 10));
    });
    const newRid = 'rId' + (maxRid + 1);
    let relsXml = meta.relsXml.replace('</Relationships>',
      '<Relationship Id="' + newRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' + newPath.split('/').pop() + '"/></Relationships>');
    zip.file(meta.relsPath, relsXml);

    let maxSheetId = 0;
    meta.sheets.forEach(function (s) { maxSheetId = Math.max(maxSheetId, s.sheetId || 0); });
    const newSheetId = maxSheetId + 1;
    const wbXml = meta.wbXml.replace('</sheets>',
      '<sheet name="' + escapeXml(newName) + '" sheetId="' + newSheetId + '" r:id="' + newRid + '"/></sheets>');
    zip.file(meta.wbPath, wbXml);

    meta.wbXml = wbXml;
    meta.relsXml = relsXml;
    meta.rels[newRid] = { id: newRid, target: 'worksheets/' + newPath.split('/').pop(), type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' };
    const info = {
      index: meta.sheets.length, name: newName, sheetId: newSheetId, state: 'visible',
      rid: newRid, path: newPath, isWorksheet: true,
    };
    meta.sheets.push(info);
    return { info: info, xml: newXml };
  }

  // ---------------------------------------------------------------- 下载

  function buildBlob(zip, onProgress) {
    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    }, function (meta) {
      if (onProgress) onProgress(meta.percent);
    });
  }

  return {
    MAX_LEVEL: MAX_LEVEL,
    CODE_RULES: CODE_RULES,
    FAMILY_PARENTS: FAMILY_PARENTS,
    FAMILY_LABELS: FAMILY_LABELS,
    norm: norm,
    codeFamily: codeFamily,
    familyLabel: familyLabel,
    colLetter: colLetter,
    parseXml: parseXml,
    readSharedStrings: readSharedStrings,
    readSheetValues: readSheetValues,
    findHeader: findHeader,
    detectLevels: detectLevels,
    detectSheetValues: detectSheetValues,
    validateSheetSums: validateSheetSums,
    toNumberValue: toNumberValue,
    readOutline: readOutline,
    outlineDist: outlineDist,
    applyOutline: applyOutline,
    makeRowMap: makeRowMap,
    maxOf: maxOf,
    applyAutoOutline: applyAutoOutline,
    applyClearOutline: applyClearOutline,
    getWorkbookMeta: getWorkbookMeta,
    loadWorkbook: loadWorkbook,
    duplicateSheet: duplicateSheet,
    makeCopyName: makeCopyName,
    buildBlob: buildBlob,
  };
});

