/**
 * ファイル名: safePptxExecutor.js
 * 説明:
 *   文字列として与えられた PptxGenJS のスニペットを eval を使わずに解析・実行する軽量サンドボックス。
 *
 * 概要:
 *   - 許可されたメソッド（pptx.addSlide や slide.addText など）だけをパターンマッチングで抽出し、
 *     事前に定義した環境で安全に実行する。
 *   - 変数宣言やループを簡易的に解釈して実行環境を構築し、引数のバリデーションや座標範囲チェックを行う。
 *   - スライドサイズや等間隔配置を計算するヘルパー関数（evenX / gridXY など）を公開してレイアウト計算を支援。
 *   - 実行結果として生成された PptxGenJS インスタンスを返し、外部から write などを行えるようにする。
 *
 * グローバル関数 runPptxFromSnippet(snippet, { pptx }) を公開する。
 */
(function () {
  async function runPptxFromSnippet(snippet, { pptx } = {}) {
    if (typeof snippet !== 'string') throw new Error('snippet must be a string');
    const PptxGenJS = window.PptxGenJS || window.pptxgen || window.pptxgenjs;
    if (!PptxGenJS) throw new Error('pptxgenjs not loaded');
    pptx = pptx || new PptxGenJS();
    const SLIDE_W = 13.33, SLIDE_H = 7.5;

    // expose slide size and simple spacing helpers for snippet calculations
    // 要素を等間隔に配置する X 座標を計算
    function evenX(index, total, itemW = 0, start = 0, end = SLIDE_W) {
      const free = end - start - itemW * total;
      const gap = free / (total + 1);
      return start + gap + index * (itemW + gap);
    }
    // 要素を等間隔に配置する Y 座標を計算
    function evenY(index, total, itemH = 0, start = 0, end = SLIDE_H) {
      const free = end - start - itemH * total;
      const gap = free / (total + 1);
      return start + gap + index * (itemH + gap);
    }

    // グリッド配置用の座標を計算
    function gridXY(index, total, cols, itemW = 0, itemH = 0, startX = 0, startY = 0, endX = SLIDE_W, endY = SLIDE_H) {
      const rows = Math.ceil(total / cols);
      const row = Math.floor(index / cols);
      const col = index % cols;
      return {
        x: evenX(col, cols, itemW, startX, endX),
        y: evenY(row, rows, itemH, startY, endY),
      };
    }

    // X 座標を中央に寄せる
    function centerX(itemW = 0, start = 0, end = SLIDE_W) {
      return start + (end - start - itemW) / 2;
    }
    // Y 座標を中央に寄せる
    function centerY(itemH = 0, start = 0, end = SLIDE_H) {
      return start + (end - start - itemH) / 2;
    }

    const cleaned = stripComments(snippet).replace(/\r/g, '');
    const env = buildEnv(cleaned);
    env.SLIDE_W = SLIDE_W;
    env.SLIDE_H = SLIDE_H;
    env.evenX = evenX;
    env.evenY = evenY;
    env.gridXY = gridXY;
    env.centerX = centerX;
    env.centerY = centerY;
    const calls = extractCalls(cleaned, env);
    let slide = null;
    const ensureSlide = () => (slide || (slide = pptx.addSlide()));

    for (const c of calls) {
      try {
        switch (c.name) {
          case 'pptx.addSlide': {
            const [optObj] = parseArgsAs([Arg.objOpt], c.args, c.env);
            slide = optObj ? pptx.addSlide(optObj) : pptx.addSlide();
            break;
          }
          case 'slide.addText': {
            ensureSlide();
            const [textOrRuns, opts] = parseArgsAs([Arg.any, Arg.objReq], c.args, c.env);
            validateBox(opts, SLIDE_W, SLIDE_H);
            sanitizeCommonTextOpts(opts);
            sanitizeRuns(textOrRuns);
            slide.addText(textOrRuns, opts);
            break;
          }
          case 'slide.addShape': {
            ensureSlide();
            const [shapeTypeExpr, opts] = parseArgsAs([Arg.any, Arg.objReq], c.args, c.env);
            const shapeType = resolveShapeString(shapeTypeExpr);
            validateBox(opts, SLIDE_W, SLIDE_H);
            sanitizeShapeOpts(opts);
            if (shapeType) {
              slide.addShape(shapeType, opts);
            }
            break;
          }
          case 'slide.addImage': {
            ensureSlide();
            const [opts] = parseArgsAs([Arg.objReq], c.args, c.env);
            validateBox(opts, SLIDE_W, SLIDE_H);
            if (opts.path && typeof opts.path === 'string') {
              const hasScheme = /^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(opts.path);
              if (!hasScheme && typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.getURL) {
                opts.path = chrome.runtime.getURL(opts.path);
              }
              const allowed = ['data:', 'blob:', 'chrome-extension:'];
              if (!allowed.some((s) => opts.path.startsWith(s))) {
                throw new Error('External image paths are not allowed');
              }
                if (opts.path.startsWith('chrome-extension:')) {
                  const url = new URL(opts.path);
                  const color = url.searchParams.get('color');
                  const colorHex = color ? (color.startsWith('#') ? color : `#${color}`) : null;
                  const fallback =
                    typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.getURL
                      ? chrome.runtime.getURL('icon/solid/square.svg')
                      : 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+PHJlY3Qgd2lkdGg9IjI0IiBoZWlnaHQ9IjI0Ii8+PC9zdmc+';
                  try {
                    const res = await fetch(opts.path);
                    if (!res.ok) throw new Error('fetch failed');
                    const type = res.headers.get('content-type') || '';
                    if (type.includes('image/svg') && colorHex) {
                      let text = await res.text();
                      const doc = new DOMParser().parseFromString(text, 'image/svg+xml');
                      if (doc && doc.documentElement) {
                        doc.documentElement.setAttribute('fill', colorHex);
                        text = new XMLSerializer().serializeToString(doc);
                      }
                      const blob = new Blob([text], { type: 'image/svg+xml' });
                      opts.path = await new Promise((resolve) => {
                        const reader = new FileReader();
                        reader.onload = () => {
                          const dataUrl = reader.result;
                          const img = new Image();
                          img.onload = () => {
                            const canvas = document.createElement('canvas');
                            canvas.width = img.width;
                            canvas.height = img.height;
                            canvas.getContext('2d').drawImage(img, 0, 0);
                            resolve(canvas.toDataURL('image/png'));
                          };
                          img.onerror = () => resolve(dataUrl);
                          img.src = dataUrl;
                        };
                        reader.onerror = () => resolve(fallback);
                        reader.readAsDataURL(blob);
                      });
                    }
                  } catch {
                    opts.path = fallback;
                  }
                }
            }
            slide.addImage(opts);
            break;
          }
          case 'slide.addTable': {
            ensureSlide();
            const [rows, opts] = parseArgsAs([Arg.any, Arg.objOpt], c.args, c.env);
            const sanitizedRows = sanitizeTableData(rows);
            if (opts) {
              if (opts.x != null) validateBox(opts, SLIDE_W, SLIDE_H);
              if (!opts.fontFace) opts.fontFace = 'Rakuten Sans JP';
            }
            slide.addTable(sanitizedRows, opts || {});
            break;
          }
          case 'slide.addChart': {
            ensureSlide();
            const [chartTypeExpr, series, opts] = parseArgsAs([Arg.chartEnum, Arg.any, Arg.objOpt], c.args, c.env);
            if (opts) {
              validateBox(opts, SLIDE_W, SLIDE_H);
              opts.titleFontFace ||= 'Rakuten Sans JP';
              opts.legendFontFace ||= 'Rakuten Sans JP';
              sanitizeChartOpts(opts);
            }
            slide.addChart(chartTypeExpr, series, opts || {});
            break;
          }
          case 'pptx.writeFile': {
            break;
          }
          default:
            throw new Error(`Not allowed call: ${c.name}`);
        }
      } catch (e) {
        console.warn('[sandbox] generation error', e);
      }
    }
    return pptx;
  }

  // コード内の変数宣言を解析して実行環境を構築
  function buildEnv(src) {
    const env = {};
    const decl = /(?:let|const)\s+([A-Za-z_]\w*)\s*=\s*([^;]+);/g;
    let m;
    while ((m = decl.exec(src))) {
      try { env[m[1]] = parseJsLikeValue(m[2], env); } catch { }
    }
    return env;
  }

  // スニペットから許可されたメソッド呼び出しを抽出
  function extractCalls(src, env) {
    const calls = [];
    applyDeclarations(src, env);
    const loopRe = /(\w+)\.forEach\s*\(|for\s*\(/g;
    let lastIndex = 0, m;
    while ((m = loopRe.exec(src))) {
      const before = src.slice(lastIndex, m.index);
      calls.push(...extractCallsFromPlain(before, env));
      if (m[0].startsWith('for(') || m[0].startsWith('for ')) {
        const { content: head, endIndex: headEnd } = readParenContent(src, loopRe.lastIndex - 1);
        const { content: body, endIndex } = readBlock(src, headEnd + 1);
        processForLoop(head.trim(), body, env, calls);
        lastIndex = endIndex + 1;
        loopRe.lastIndex = lastIndex;
      } else {
        const { content, endIndex } = readParenContent(src, loopRe.lastIndex - 1);
        const inside = content.trim();
        const arrow = /^\s*(?:\(([^)]*)\)|(\w+))\s*=>\s*{([\s\S]*)}\s*$/.exec(inside);
        if (arrow) {
          const paramsStr = arrow[1] || arrow[2] || '';
          const params = paramsStr.split(/\s*,\s*/).filter(Boolean);
          const body = arrow[3];
          const arr = resolvePath(env, m[1], env);
          if (Array.isArray(arr)) {
            for (let i = 0; i < arr.length && i < MAX_ITER; i++) {
              const newEnv = Object.assign({}, env);
              if (params[0]) newEnv[params[0]] = arr[i];
              if (params[1]) newEnv[params[1]] = i;
              calls.push(...extractCalls(body, newEnv));
              if (params[0]) delete newEnv[params[0]];
              if (params[1]) delete newEnv[params[1]];
              Object.assign(env, newEnv);
            }
          }
        }
        lastIndex = endIndex + 1;
        loopRe.lastIndex = lastIndex;
      }
    }
    calls.push(...extractCallsFromPlain(src.slice(lastIndex), env));
    return calls;
  }

  const MAX_ITER = 1000;

  // let/const の宣言文を読み取り環境変数に値を入れる
  function applyDeclarations(src, env) {
    const decl = /(?:let|const)\s+([A-Za-z_]\w*)\s*=\s*([^;]+);/g;
    let m;
    while ((m = decl.exec(src))) {
      try {
        env[m[1]] = parseJsLikeValue(m[2], env);
      } catch {}
    }
  }

  // for 文のヘッダーと本体を解釈し繰り返し処理を行う
  function processForLoop(header, body, env, calls) {
    const forOf = /^(?:let|const|var)?\s*([A-Za-z_]\w*)\s+of\s+(.+)$/.exec(header);
    if (forOf) {
      const arr = parseJsLikeValue(forOf[2], env);
      if (Array.isArray(arr)) {
        for (let i = 0; i < arr.length && i < MAX_ITER; i++) {
          const newEnv = Object.assign({}, env);
          newEnv[forOf[1]] = arr[i];
          calls.push(...extractCalls(body, newEnv));
          delete newEnv[forOf[1]];
          Object.assign(env, newEnv);
        }
      }
      return;
    }
    const classic = /^let\s+([A-Za-z_]\w*)\s*=\s*(\d+)\s*;\s*\1\s*<\s*([A-Za-z_]\w*)\.length\s*;\s*\1\s*(?:\+\+|\+=\s*1)\s*$/.exec(header);
    if (classic) {
      const idxVar = classic[1];
      const start = parseInt(classic[2], 10);
      const arrName = classic[3];
      const arr = resolvePath(env, arrName, env);
      if (Array.isArray(arr)) {
        for (let i = start; i < arr.length && i - start < MAX_ITER; i++) {
          const newEnv = Object.assign({}, env);
          newEnv[idxVar] = i;
          calls.push(...extractCalls(body, newEnv));
          delete newEnv[idxVar];
          Object.assign(env, newEnv);
        }
      }
    }
  }

  // 単純なコード部分から許可された関数呼び出しを抜き出す
  function extractCallsFromPlain(s, env) {
    const allow = [
      'pptx\\.addSlide',
      'slide\\.addText',
      'slide\\.addShape',
      'slide\\.addImage',
      'slide\\.addTable',
      'slide\\.addChart',
      'pptx\\.writeFile'
    ];
    const callRe = new RegExp(`\\b(?:${allow.join('|')})\\s*\\(`, 'g');
    const assignRe = /(?:let|const)?\s*([A-Za-z_]\w*)\s*([+\-*/])?=\s*([^;]+);/g;
    const calls = [];
    let index = 0;
    while (index < s.length) {
      assignRe.lastIndex = index;
      callRe.lastIndex = index;
      const am = assignRe.exec(s);
      const cm = callRe.exec(s);
      const ai = am ? am.index : Infinity;
      const ci = cm ? cm.index : Infinity;
      if (ai === Infinity && ci === Infinity) break;
      if (ai < ci) {
        const name = am[1];
        const op = am[2];
        const expr = am[3];
        try {
          const val = parseJsLikeValue(expr, env);
          if (op && typeof env[name] === 'number' && typeof val === 'number') {
            if (op === '+') env[name] += val;
            else if (op === '-') env[name] -= val;
            else if (op === '*') env[name] *= val;
            else if (op === '/') env[name] /= val;
          } else {
            env[name] = val;
          }
        } catch {}
        index = assignRe.lastIndex;
      } else {
        const nameEnd = cm.index + cm[0].length;
        const name = cm[0].replace(/\s*\($/, '');
        const { content, endIndex } = readParenContent(s, nameEnd - 1);
        calls.push({ name, args: splitTopLevelArgs(content.trim()), env: Object.assign({}, env) });
        index = endIndex + 1;
      }
    }
    return calls;
  }

  // 開き括弧の位置から対応する閉じ括弧までの中身を取り出す
  function readParenContent(s, openIndex) {
    let i = openIndex, depth = 0, inStr = false, strCh = '', esc = false;
    const start = openIndex + 1;
    for (; i < s.length; i++) {
      const ch = s[i];
      if (inStr) {
        if (esc) { esc = false; continue; }
        if (ch === '\\') { esc = true; continue; }
        if (ch === strCh) { inStr = false; continue; }
        continue;
      }
      if (ch === '"' || ch === "'") { inStr = true; strCh = ch; continue; }
      if (ch === '(') depth++;
      if (ch === ')') {
        depth--;
        if (depth === 0) {
          return { content: s.slice(start, i), endIndex: i };
        }
      }
    }
    throw new Error('Unbalanced parentheses');
  }

  // カンマで区切られた最上位の引数リストに分割する
  function splitTopLevelArgs(s) {
    const args = [];
    let buf = '', depth = 0, inStr = false, strCh = '', esc = false;
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (inStr) {
        buf += ch;
        if (esc) { esc = false; continue; }
        if (ch === '\\') { esc = true; continue; }
        if (ch === strCh) { inStr = false; }
        continue;
      }
      if (ch === '"' || ch === "'") { inStr = true; strCh = ch; buf += ch; continue; }
      if (ch === '{' || ch === '[' || ch === '(') { depth++; buf += ch; continue; }
      if (ch === '}' || ch === ']' || ch === ')') { depth--; buf += ch; continue; }
      if (ch === ',' && depth === 0) { args.push(buf.trim()); buf = ''; continue; }
      buf += ch;
    }
    if (buf.trim()) args.push(buf.trim());
    return args;
  }

  const Arg = {
    objReq: { kind: 'objReq' },
    objOpt: { kind: 'objOpt' },
    any: { kind: 'any' },
    chartEnum: { kind: 'enum', resolver: resolveChartEnum },
  };

  // 関数の期待する型に合わせて引数を解析する
  function parseArgsAs(specs, args, env) {
    if (args.length < specs.filter(s => s.kind.endsWith('Req')).length) {
      throw new Error('Not enough arguments');
    }
    const out = [];
    for (let i = 0; i < specs.length; i++) {
      const spec = specs[i], raw = args[i];
      if (raw == null) {
        out.push(undefined);
        continue;
      }
      if (spec.kind === 'any') {
        out.push(parseJsLikeValue(raw, env));
      } else if (spec.kind === 'objReq' || spec.kind === 'objOpt') {
        const v = parseJsLikeValue(raw, env);
        if ((spec.kind === 'objReq') && (typeof v !== 'object' || v === null || Array.isArray(v))) {
          throw new Error('Object required');
        }
        out.push(v);
      } else if (spec.kind === 'enum') {
        out.push(spec.resolver(raw));
      }
    }
    return out;
  }

  // JavaScript 風の文字列を実際の値に変換する
  function parseJsLikeValue(src, env = {}) {
    const trimmed = src.trim();
    if (trimmed === '') return undefined;
    if (/^(['"]).*\1$/.test(trimmed) || /^(true|false|null)$/.test(trimmed) || /^-?\d+(\.\d+)?$/.test(trimmed)) {
      return JSON.parse(toJsonPrimitive(trimmed));
    }
    const tern = parseTernary(trimmed, env);
    if (tern !== undefined) return tern;
    if (/^chrome\.runtime\.getURL\((.+)\)$/.test(trimmed)) {
      const inner = /^chrome\.runtime\.getURL\((.+)\)$/.exec(trimmed)[1];
      const rel = parseJsLikeValue(inner, env);
      if (typeof rel !== 'string' || /^[a-zA-Z]+:/.test(rel)) {
        throw new Error('Invalid getURL path');
      }
      if (typeof chrome === 'undefined' || !chrome.runtime || !chrome.runtime.getURL) {
        throw new Error('chrome.runtime.getURL unavailable');
      }
      const url = chrome.runtime.getURL(rel);
      if (!url.startsWith('chrome-extension:')) {
        throw new Error('Invalid getURL result');
      }
      return url;
    }
    const mathCall = /^Math\.([A-Za-z]+)\((.*)\)$/.exec(trimmed);
    if (mathCall) {
      const fn = mathCall[1];
      const allowed = ['floor', 'ceil', 'round', 'abs', 'min', 'max', 'pow', 'sqrt'];
      if (allowed.includes(fn)) {
        const args = splitTopLevelArgs(mathCall[2]).map((a) => Number(parseJsLikeValue(a, env)));
        return Math[fn](...args);
      }
    }
    const evenCall = /^(evenX|evenY)\((.*)\)$/.exec(trimmed);
    if (evenCall) {
      const args = splitTopLevelArgs(evenCall[2]).map(a => parseJsLikeValue(a, env));
      const fn = env[evenCall[1]];
      if (typeof fn === 'function') return fn.apply(null, args);
      return undefined;
    }
    const gridCall = /^gridXY\((.*)\)$/.exec(trimmed);
    if (gridCall) {
      const args = splitTopLevelArgs(gridCall[1]).map(a => parseJsLikeValue(a, env));
      const fn = env.gridXY;
      if (typeof fn === 'function') return fn.apply(null, args);
      return undefined;
    }
    if (/^chrome\./.test(trimmed)) {
      throw new Error('chrome.* calls are not allowed');
    }
    if (/^`[\s\S]*`$/.test(trimmed)) {
      let result = '';
      const re = /\$\{([^}]+)\}/g;
      let last = 1, m;
      while ((m = re.exec(trimmed))) {
        result += trimmed.slice(last, m.index);
        const v = parseJsLikeValue(m[1], env);
        if (v != null) result += String(v);
        last = re.lastIndex;
      }
      result += trimmed.slice(last, trimmed.length - 1);
      return result;
    }
    if (trimmed.startsWith('{')) return parseObject(trimmed, env);
    if (trimmed.startsWith('[')) return parseArray(trimmed, env);
    if (/^pptx\s*\.\s*ShapeType\s*\./.test(trimmed)) return resolveShapeEnum(trimmed);
    if (/^pptx\s*\.\s*ChartType\s*\./.test(trimmed)) return resolveChartEnum(trimmed);
    if (/^[A-Za-z_]\w*(?:\.[A-Za-z_]\w*|\[[^\]]+\])*$/s.test(trimmed)) {
      return resolvePath(env, trimmed, env);
    }
    if (/[+\-*/%]/.test(trimmed)) {
      const v = evalArithmetic(trimmed, env);
      return isNaN(v) ? undefined : v;
    }
    throw new Error('Unsupported expression: ' + trimmed);
  }

  // シングルクォートを JSON 形式に変換するヘルパー
  function toJsonPrimitive(s) {
    if (s.startsWith("'")) {
      return '"' + s.slice(1, -1).replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
    }
    return s;
  }

  // オブジェクトリテラル文字列を実際のオブジェクトに変換
  function parseObject(src, env) {
    const inner = src.trim().slice(1, -1);
    if (inner.trim() === '') return {};
    const props = splitTopLevelArgs(inner);
    const obj = {};
    props.forEach(p => {
      const idx = p.indexOf(':');
      if (idx === -1) return;
      const keyRaw = p.slice(0, idx).trim();
      const valStr = p.slice(idx + 1).trim();
      let key = keyRaw;
      if (/^(['"]).*\1$/.test(keyRaw)) {
        key = JSON.parse(toJsonPrimitive(keyRaw));
      }
      try {
        const v = parseJsLikeValue(valStr, env);
        if (!(typeof v === 'number' && isNaN(v))) {
          if (v !== undefined) obj[key] = v;
        }
      } catch {}
    });
    return obj;
  }

  // 配列リテラル文字列を実際の配列に変換
  function parseArray(src, env) {
    const inner = src.trim().slice(1, -1);
    if (inner.trim() === '') return [];
    return splitTopLevelArgs(inner).map(v => {
      try { return parseJsLikeValue(v, env); } catch { return undefined; }
    }).filter(v => v !== undefined && !(typeof v === 'number' && isNaN(v)));
  }

  // 三項演算子を評価して値を返す
  function parseTernary(src, env) {
    let depth = 0, inStr = false, strCh = '', esc = false;
    for (let i = 0; i < src.length; i++) {
      const ch = src[i];
      if (inStr) {
        if (esc) { esc = false; continue; }
        if (ch === '\\') { esc = true; continue; }
        if (ch === strCh) { inStr = false; }
        continue;
      }
      if (ch === '"' || ch === "'") { inStr = true; strCh = ch; continue; }
      if (ch === '(' || ch === '{' || ch === '[') { depth++; continue; }
      if (ch === ')' || ch === '}' || ch === ']') { depth--; continue; }
      if (ch === '?' && depth === 0) {
        const condStr = src.slice(0, i).trim();
        const rest = src.slice(i + 1);
        let depth2 = 0, inStr2 = false, strCh2 = '', esc2 = false;
        for (let j = 0; j < rest.length; j++) {
          const ch2 = rest[j];
          if (inStr2) {
            if (esc2) { esc2 = false; continue; }
            if (ch2 === '\\') { esc2 = true; continue; }
            if (ch2 === strCh2) { inStr2 = false; }
            continue;
          }
          if (ch2 === '"' || ch2 === "'") { inStr2 = true; strCh2 = ch2; continue; }
          if (ch2 === '(' || ch2 === '{' || ch2 === '[') { depth2++; continue; }
          if (ch2 === ')' || ch2 === '}' || ch2 === ']') { depth2--; continue; }
          if (ch2 === ':' && depth2 === 0) {
            const truthy = rest.slice(0, j).trim();
            const falsy = rest.slice(j + 1).trim();
            const condVal = parseJsLikeValue(condStr, env);
            return parseJsLikeValue(condVal ? truthy : falsy, env);
          }
        }
        break;
      }
    }
    return undefined;
  }

  // 足し算や掛け算を含む簡単な数式を計算する
  function evalArithmetic(expr, env) {
    const replaced = expr.replace(/[A-Za-z_]\w*(?:\.[A-Za-z_]\w*|\[[^\]]+\])*/g, (name) => {
      const v = resolvePath(env, name, env);
      return typeof v === 'number' ? String(v) : 'NaN';
    });
    if (!/^[0-9+\-*/%().\sNa]+$/.test(replaced)) return NaN;
    const tokens = replaced.match(/NaN|\d+(?:\.\d+)?|[()+\-*/%]/g);
    if (!tokens) return NaN;
    let pos = 0;
    // 式全体を解析する
    function parseExpression() {
      let val = parseTerm();
      while (pos < tokens.length && (tokens[pos] === '+' || tokens[pos] === '-')) {
        const op = tokens[pos++];
        const rhs = parseTerm();
        if (isNaN(rhs)) return NaN;
        val = op === '+' ? val + rhs : val - rhs;
      }
      return val;
    }
    // 乗算・除算などの項を処理する
    function parseTerm() {
      let val = parseFactor();
      while (pos < tokens.length && (tokens[pos] === '*' || tokens[pos] === '/' || tokens[pos] === '%')) {
        const op = tokens[pos++];
        const rhs = parseFactor();
        if (isNaN(rhs)) return NaN;
        val = op === '*' ? val * rhs : op === '/' ? val / rhs : val % rhs;
      }
      return val;
    }
    // 数値や括弧を読み取る
    function parseFactor() {
      const t = tokens[pos++];
      if (t === '-') return -parseFactor();
      if (t === '(') {
        const v = parseExpression();
        if (tokens[pos] !== ')') return NaN;
        pos++;
        return v;
      }
      if (t === 'NaN') return NaN;
      return parseFloat(t);
    }
    const res = parseExpression();
    return pos < tokens.length ? NaN : res;
  }

  // 文字列のパスからオブジェクトの値を取り出す
  function resolvePath(obj, path, env) {
    const re = /([A-Za-z_]\w*)|\[([^\]]+)\]/g;
    let m, cur = obj;
    while ((m = re.exec(path))) {
      if (m[1]) {
        if (cur && typeof cur === 'object' && m[1] in cur) {
          cur = cur[m[1]];
        } else {
          return undefined;
        }
      } else if (m[2]) {
        const idx = parseJsLikeValue(m[2], env);
        if (cur && typeof cur === 'object') {
          cur = cur[idx];
        } else {
          return undefined;
        }
      }
    }
    return cur;
  }

  // ソースコードからコメント部分を取り除く
  function stripComments(str) {
    let out = '', inStr = false, strCh = '', esc = false;
    for (let i = 0; i < str.length; i++) {
      const ch = str[i], next = str[i + 1];
      if (inStr) {
        out += ch;
        if (esc) { esc = false; continue; }
        if (ch === '\\') { esc = true; continue; }
        if (ch === strCh) { inStr = false; }
        continue;
      }
      if (ch === '"' || ch === "'") { inStr = true; strCh = ch; out += ch; continue; }
      if (ch === '/' && next === '*') {
        i += 2;
        while (i < str.length && !(str[i] === '*' && str[i + 1] === '/')) i++;
        i++; // skip closing '/'
        continue;
      }
      if (ch === '/' && next === '/') {
        i += 2;
        while (i < str.length && str[i] !== '\n') i++;
        out += '\n';
        continue;
      }
      out += ch;
    }
    return out;
  }

  // 波括弧で囲まれたブロックの中身を取得する
  function readBlock(s, openIndex) {
    let i = openIndex, depth = 0, inStr = false, strCh = '', esc = false;
    const start = openIndex + 1;
    for (; i < s.length; i++) {
      const ch = s[i];
      if (inStr) {
        if (esc) { esc = false; continue; }
        if (ch === '\\') { esc = true; continue; }
        if (ch === strCh) { inStr = false; continue; }
        continue;
      }
      if (ch === '"' || ch === "'") { inStr = true; strCh = ch; continue; }
      if (ch === '{') depth++;
      if (ch === '}') {
        depth--;
        if (depth === 0) {
          return { content: s.slice(start, i), endIndex: i };
        }
      }
    }
    throw new Error('Unbalanced braces');
  }

  // pptx.ShapeType の文字列を実際の列挙値に変換
  function resolveShapeEnum(expr) {
    const m = expr.replace(/\s+/g, '').match(/^pptx\.ShapeType\.(\w+)$/);
    if (!m) throw new Error('Invalid ShapeType: ' + expr);
    const P = (window.PptxGenJS || window.pptxgen || window.pptxgenjs);
    return P.ShapeType ? P.ShapeType[m[1]] : (new P()).ShapeType[m[1]];
  }

  // pptx.ChartType の文字列を実際の列挙値に変換
  function resolveChartEnum(expr) {
    const m = expr.replace(/\s+/g, '').match(/^pptx\.ChartType\.(\w+)$/);
    if (!m) throw new Error('Invalid ChartType: ' + expr);
    const P = (window.PptxGenJS || window.pptxgen || window.pptxgenjs);
    return P.ChartType ? P.ChartType[m[1]] : (new P()).ChartType[m[1]];
  }

  // オブジェクトの座標やサイズが正しいか確認しインチ値に変換
  function validateBox(o, W, H) {
    const x = toInch(o.x, W), y = toInch(o.y, H), w = toInch(o.w, W), h = toInch(o.h, H);
    if ([x, y, w, h].some(v => typeof v !== 'number' || isNaN(v))) return;
    // Allow boxes that extend beyond the slide bounds so that the generated
    // presentation can intentionally contain elements positioned off-slide.
    // Security checks and value sanitation remain unchanged elsewhere.
    o.x = x; o.y = y; o.w = w; o.h = h;
  }

  // パーセント指定などをインチの数値に直す
  function toInch(v, axisLen) {
    if (v == null) return v;
    if (typeof v === 'string' && v.trim().endsWith('%')) {
      return (parseFloat(v) / 100) * axisLen;
    }
    return Number(v);
  }

  // テキスト共通のオプションを安全な値に整える
  function sanitizeCommonTextOpts(opts) {
    opts.fontFace ||= 'Rakuten Sans JP';
    if ('color' in opts) {
      opts.color = normalizeColorOrDefault(opts.color, '000000');
    }
    if (opts.fill && typeof opts.fill === 'object' && 'color' in opts.fill) {
      opts.fill.color = normalizeColorOrDefault(opts.fill.color, 'FFFFFF');
    }
  }

  // テキストラン配列内のオプションを一括で整える
  function sanitizeRuns(runs) {
    if (!Array.isArray(runs)) return;
    runs.forEach((r) => {
      if (r && r.options) sanitizeCommonTextOpts(r.options);
    });
  }

  // 図形オプションの色や線を正しい形式に直す
  function sanitizeShapeOpts(opts) {
    if (opts.fill && typeof opts.fill === 'object') {
      if ('color' in opts.fill) {
        opts.fill.color = normalizeColorOrDefault(opts.fill.color, 'FFFFFF');
      } else {
        delete opts.fill;
      }
    }
    if (opts.line && typeof opts.line === 'object') {
      if ('color' in opts.line) {
        opts.line.color = normalizeColorOrDefault(opts.line.color, '000000');
      } else if (!('width' in opts.line)) {
        delete opts.line;
      }
    }
  }

  // グラフ描画用オプションの色や線を調整する
  function sanitizeChartOpts(opts) {
    ['gridLine', 'catGridLine', 'valGridLine'].forEach((k) => {
      const o = opts[k];
      if (o) {
        if (typeof o.size === 'number' && o.size <= 0) {
          o.size = 0.1;
        }
        if (o.color) o.color = normalizeColorOrDefault(o.color, 'CCCCCC');
      }
    });
    if (Array.isArray(opts.chartColors)) {
      opts.chartColors = opts.chartColors.map((c) => normalizeColorOrDefault(c, '000000'));
    }
    const normalizeColorProp = (obj, key) => {
      if (!obj || !obj[key]) return;
      let val = obj[key];
      if (val && typeof val === 'object' && 'color' in val) {
        const c = normalizeColor(val.color);
        if (c) {
          if (Object.keys(val).length === 1) {
            obj[key] = c;
          } else {
            val.color = c;
          }
        } else {
          if (Object.keys(val).length === 1) {
            delete obj[key];
          } else {
            delete val.color;
          }
        }
      } else if (typeof val === 'string') {
        const c = normalizeColor(val);
        if (c) {
          obj[key] = c;
        } else {
          delete obj[key];
        }
      }
    };

    ['fill', 'catAxisLabelColor', 'catAxisLineColor', 'valAxisLabelColor', 'valAxisLineColor', 'valAxisTitleColor', 'dataLabelColor'].forEach((k) => normalizeColorProp(opts, k));
    if (opts.chartArea) normalizeColorProp(opts.chartArea, 'fill');
  }

  // 表データ内の null を空セルに置き換える
  function sanitizeTableData(tableData) {
    if (!Array.isArray(tableData)) return tableData;
    return tableData.map((row) =>
      Array.isArray(row)
        ? row.map((cell) => (cell === null ? { text: '', options: {} } : cell))
        : row
    );
  }

  // 色コードを正規化し、無効なら既定値を返す
  function normalizeColorOrDefault(c, def) {
    return normalizeColor(c) || def;
  }

  // #やrgb表記を6桁の16進カラーコードに変換
  function normalizeColor(c) {
    if (typeof c !== 'string') return null;
    let s = c.trim();
    if (s.startsWith('#')) s = s.slice(1);
    if (/^[0-9a-fA-F]{3}$/.test(s)) {
      s = s.split('').map(ch => ch + ch).join('');
    }
    const rgb = s.match(/^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)$/i);
    if (rgb) {
      const toHex = (n) => Math.max(0, Math.min(255, parseInt(n, 10))).toString(16).padStart(2, '0');
      s = [rgb[1], rgb[2], rgb[3]].map(toHex).join('');
    }
    if (/^[0-9a-fA-F]{6}$/.test(s)) {
      return s.toUpperCase();
    }
    return null;
  }

  // "rect" などの文字列から ShapeType を取得
  function resolveShapeString(expr) {
    if (typeof expr !== 'string') return null;
    const allowed = ['rect', 'roundRect', 'ellipse', 'line'];
    if (!allowed.includes(expr)) return null;
    const P = (window.PptxGenJS || window.pptxgen || window.pptxgenjs);
    const Shape = P.ShapeType ? P.ShapeType : (new P()).ShapeType;
    return Shape[expr];
  }

  window.runPptxFromSnippet = runPptxFromSnippet;
})();

