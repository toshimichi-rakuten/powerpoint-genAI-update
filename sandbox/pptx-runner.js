/**
 * ファイル名: sandbox/pptx-runner.js
 * 説明:
 *   拡張機能とは分離された sandbox iframe 内で実行され、
 *   渡された PptxGenJS スニペットを安全に走らせて PPTX データを親へ返すランナー。
 *
 * 主な機能:
 *   - window/chrome が存在しない環境でのポリフィル定義と、拡張内リソースのみを許可する fetch の上書き。
 *   - 進捗や完了状態を postMessage で親へ通知するための safePostMessage。
 *   - アイコン画像の DataURL 化と欠損時フォールバック、マスター スライドの定義、画像座標の EMU→インチ変換など PPTX 生成処理。
 *   - `compression: 'STORE'` で ZIP 圧縮を省略しつつ、生成後は結果を dataURL として送り返す。
 */

'use strict';

// --- sandbox 実行用ポリフィル ---
if (typeof window === 'undefined') {
  global.window = { addEventListener: () => {}, parent: { postMessage: () => {} } };
}
if (typeof chrome === 'undefined' || !chrome.runtime || !chrome.runtime.getURL) {
  const base = typeof location === 'object'
    ? new URL('../', location.href).href
    : '';
  window.chrome = { runtime: { getURL: (path) => base + path } };
}

let parentOrigin = (() => {
  try {
    return new URL(document.referrer).origin;
  } catch (e) {
    return null;
  }
})();

// 親フレームにメッセージを安全に送る
function safePostMessage(message) {
  if (!parentOrigin) {
    console.error('[sandbox] unknown parent origin');
    return;
  }
  window.parent.postMessage(message, parentOrigin);
}

const originalFetch = window.fetch;
// 拡張機能内のファイルだけを取得できるよう fetch を上書き
window.fetch = function(input, init) {
  const url = typeof input === 'string' ? input : input && input.url;
  const base = (chrome && chrome.runtime && chrome.runtime.getURL)
    ? new URL(chrome.runtime.getURL('')).origin
    : (typeof location === 'object' ? location.origin : '');
  const resolved = new URL(url, base);
  if (resolved.origin !== base) {
    throw new Error('Network access to external resources is blocked');
  }
  return originalFetch(resolved.href, init);
};

// --- 定数定義 ---
const LAYOUT_WIDTH = 13.33;
const LAYOUT_HEIGHT = 7.5;
const MASTER_SLIDE_TITLE = 'CORPORATE_MASTER_FINAL';
const DEFAULT_FONT_FACE = 'Rakuten Sans JP';
const FALLBACK_ICON_PATH = 'icon/solid/square.svg';
const INQUIRY_URL = 'https://forms.office.com/r/04iMPYrVmv';
const MANUAL_URL = 'https://rak.box.com/s/4dvyt5n5vbvq1lbdwo36n8hdqf527qg1';

window.LAYOUT_WIDTH = LAYOUT_WIDTH;
window.LAYOUT_HEIGHT = LAYOUT_HEIGHT;
window.MASTER_SLIDE_TITLE = MASTER_SLIDE_TITLE;
window.DEFAULT_FONT_FACE = DEFAULT_FONT_FACE;
window.defineMasterSlide = defineMasterSlide;

const iconCache = new Map();

// アイコン画像を読み込み DataURL に変換
async function loadIcon(path, colorHex) {
  const cacheKey = `${path}:${colorHex || ''}`;
  if (iconCache.has(cacheKey)) {
    return iconCache.get(cacheKey);
  }
  const res = await fetch(chrome.runtime.getURL(path));
  if (!res.ok) throw new Error(`status ${res.status}`);
  const type = res.headers.get('content-type') || '';
  let blob;
  if (type.includes('image/svg')) {
    let text = await res.text();
    if (colorHex) {
      const doc = new DOMParser().parseFromString(text, 'image/svg+xml');
      if (doc && doc.documentElement) {
        doc.documentElement.setAttribute('fill', colorHex);
        text = new XMLSerializer().serializeToString(doc);
      }
    }
    blob = new Blob([text], { type: 'image/svg+xml' });
  } else {
    blob = await res.blob();
  }
  const result = await new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = reader.result;
      if (blob.type === 'image/svg+xml') {
        const img = new Image();
        img.onload = () => {
          const canvas = document.createElement('canvas');
          canvas.width = img.width;
          canvas.height = img.height;
          canvas.getContext('2d').drawImage(img, 0, 0);
          resolve(`'${canvas.toDataURL('image/png')}'`);
        };
        img.onerror = () => resolve(`'${dataUrl}'`);
        img.src = dataUrl;
      } else {
        resolve(`'${dataUrl}'`);
      }
    };
    reader.onerror = () => resolve(null);
    reader.readAsDataURL(blob);
  });
  if (result) iconCache.set(cacheKey, result);
  return result;
}

// アイコン取得に失敗したときフォールバックを適用
async function processIcon(path, colorHex) {
  const data = await loadIcon(path, colorHex).catch(() => null);
  if (data) return data;
  return (await loadIcon(FALLBACK_ICON_PATH, colorHex).catch(() => null))
    || `'${chrome.runtime.getURL(FALLBACK_ICON_PATH)}'`;
}

// EMU単位をインチに変換
function emuToInches(emu) {
  return emu / 914400;
}

// 共通デザインのマスタースライドを定義
function defineMasterSlide(pptx, logoPath) {
  try {
    const footerTextStyle = {
      fontFace: DEFAULT_FONT_FACE,
      fontSize: 9,
      color: '000000',
      bold: true,
      valign: 'middle'
    };

    pptx.defineSlideMaster({
      title: MASTER_SLIDE_TITLE,
      background: { color: 'FFFFFF' },
      objects: [
        ...(logoPath
          ? [{
              image: {
                path: logoPath,
                x: emuToInches(334965),
                y: emuToInches(6372225),
                w: emuToInches(358776),
                h: emuToInches(358776)
              }
            }]
          : []),
        {
          text: {
            text: 'CONFIDENTIAL',
            options: {
              ...footerTextStyle,
              x: 11.39915113735783,
              y: 7.0438593350831145,
              w: 1.2967629046369205,
              h: 0.2524409448818898,
              align: 'left'
            }
          }
        }
      ],
      slideNumber: {
        ...footerTextStyle,
        x: 12.671220253718285,
        y: 7.0438593350831145,
        w: 0.5691130796150481,
        h: 0.2524409448818898,
        align: 'left'
      }
    });
    return true;
  } catch (error) {
    console.error('スライドマスターの定義中にエラーが発生しました', error);
    return false;
  }
}

// --- chrome.runtime.getURL を文字列に置き換える ---
// コード内の getURL 呼び出しを実際のパスに置き換える
async function replaceGetUrls(code) {
  if (typeof chrome === 'undefined' || !chrome.runtime || !chrome.runtime.getURL) {
    return code;
  }
  const regex = /chrome\.runtime\.getURL\(\s*(['"])([^'"\)]+)\1\s*\)/g;
  const matches = [...code.matchAll(regex)];
  const promises = matches.map(async (m) => {
    const original = m[2];
    const [path, query] = original.split('?');
    const params = new URLSearchParams(query || '');
    const color = params.get('color');
    const colorHex = color ? (color.startsWith('#') ? color : `#${color}`) : null;
    return await processIcon(path, colorHex);
  });
  const results = await Promise.all(promises);
  let replaced = code;
  matches.forEach((m, idx) => {
    replaced = replaced.replace(m[0], results[idx]);
  });
  return replaced;
}

// 危険な文字列を除去してから実行
function sanitizeSnippet(code) {
  return (code || '')
    .split('\n')
    .filter(line => {
      const t = line.trim();
      if (/new\s+PptxGenJS\b/.test(t)) return false;
      if (/pptx\.write(File)?\b/.test(t)) return false;
      return true;
    })
    .join('\n');
}

// 危険な構文が含まれていないか簡易チェック
function isSafeCode(code) {
  const acorn = (typeof globalThis !== 'undefined' && globalThis.acorn) || null;
  if (!acorn) {
    // fall back to regex-based checks when parser isn't available
    const patterns = [
      /fetch\s*\(/i,
      /XMLHttpRequest/i,
      /localStorage/i,
      /chrome\.storage/i,
      /document\.cookie/i,
      /navigator\.clipboard/i,
      /\beval\b/i,
      /new\s+Function/i
    ];
    return !patterns.some((re) => re.test(code || ''));
  }
  try {
    const ast = acorn.parse(code || '', { ecmaVersion: 2020 });
    const banned = [
      'fetch',
      'XMLHttpRequest',
      'localStorage',
      'chrome',
      'document',
      'navigator',
      'eval',
      'Function'
    ];
    const allowedCalls = new Set([
      'pptx.addSlide',
      'slide.addText',
      'slide.addShape',
      'slide.addImage',
      'slide.addTable',
      'slide.addChart',
      'pptx.writeFile'
    ]);
    let safe = true;
    walkAst(ast, (node) => {
      if (!safe) return;
      if (node.type === 'Identifier') {
        if (banned.includes(node.name)) safe = false;
      } else if (node.type === 'MemberExpression') {
        const name = memberName(node);
        if (banned.some((b) => name === b || name.startsWith(b + '.'))) safe = false;
      } else if (node.type === 'CallExpression') {
        const name = memberName(node.callee);
        if (!name) { safe = false; return; }
        if (name.endsWith('.forEach')) return; // allow Array#forEach
        if (!allowedCalls.has(name)) safe = false;
      }
    });
    return safe;
  } catch (e) {
    console.warn('[isSafeCode] parse error', e);
    return false;
  }
}

// 抽象構文木を深さ優先で巡回
function walkAst(node, fn) {
  if (!node || typeof node.type !== 'string') return;
  fn(node);
  for (const key of Object.keys(node)) {
    const child = node[key];
    if (Array.isArray(child)) {
      child.forEach((c) => walkAst(c, fn));
    } else if (child && typeof child.type === 'string') {
      walkAst(child, fn);
    }
  }
}

// ASTノードからメンバー名を取得
function memberName(node) {
  if (!node) return '';
  if (node.type === 'Identifier') return node.name;
  if (node.type === 'MemberExpression' && !node.computed) {
    const obj = memberName(node.object);
    const prop = memberName(node.property);
    return obj && prop ? obj + '.' + prop : '';
  }
  return '';
}

// 色指定がオブジェクト形式なら文字列に変換
function ensureColorStrings(obj) {
  if (!obj || typeof obj !== 'object') return;
  if (Array.isArray(obj)) {
    obj.forEach((item, idx) => {
      if (item && typeof item === 'object') {
        ensureColorStrings(item);
      } else if (item != null && typeof item !== 'string') {
        obj[idx] = String(item);
      }
    });
    return;
  }
  Object.keys(obj).forEach((k) => {
    const v = obj[k];
    if (k.toLowerCase().includes('color')) {
      if (Array.isArray(v)) {
        v.forEach((item, idx) => {
          if (item && typeof item === 'object') {
            ensureColorStrings(item);
          } else if (item != null && typeof item !== 'string') {
            v[idx] = String(item);
          }
        });
      } else if (v != null && typeof v !== 'string') {
        obj[k] = String(v);
      }
    } else if (v && typeof v === 'object') {
      ensureColorStrings(v);
    }
  });
}

// すべての図形の色表記を統一
function normalizeColors(pptx) {
  const slides = pptx._slides || pptx.slides || [];
  slides.forEach((s) => {
    const objs = s._slideObjects || [];
    objs.forEach((o) => {
      if (o.options) ensureColorStrings(o.options);
      if (o.text && o.text.options) ensureColorStrings(o.text.options);
    });
  });
}

// オブジェクト内のパスを置換
function replacePaths(obj, from, to) {
  if (!obj || typeof obj !== 'object') return;
  if (Array.isArray(obj)) {
    obj.forEach((item) => replacePaths(item, from, to));
    return;
  }
  Object.keys(obj).forEach((k) => {
    const v = obj[k];
    if (typeof v === 'string') {
      if (v === from) obj[k] = to;
    } else if (v && typeof v === 'object') {
      replacePaths(v, from, to);
    }
  });
}

// writeが失敗したとき dataURL で取得する
async function writeWithFallback(pptx, options) {
  const tried = new Set();
  while (true) {
    try {
      return await pptx.write('blob', {
        compression: options.compression || 'STORE'
      });
    } catch (err) {
      const msg = err && err.message;
      const match = msg && msg.match(/Unable to load image[^:]*:\s*(.+)/);
      if (match) {
        const missing = match[1].trim();
        if (tried.has(missing)) throw err;
        tried.add(missing);
        const fallback = chrome.runtime && chrome.runtime.getURL
          ? chrome.runtime.getURL(FALLBACK_ICON_PATH)
          : FALLBACK_ICON_PATH;
        console.warn('[sandbox] replacing missing image', missing, 'with', fallback);
        replacePaths(pptx, missing, fallback);
        continue;
      }
      throw err;
    }
  }
}

window.addEventListener('message', async (e) => {
  if (e.source !== window.parent) return;
  const allowed = ['generate', 'generate-single', 'generate-multi'];
  if (!parentOrigin && allowed.includes((e.data || {}).action)) {
    parentOrigin = e.origin;
  }
  if (e.origin !== parentOrigin) return;
  const { action, code, codes = [], fileName = 'presentation.pptx', options = {} } = e.data || {};
  if (!allowed.includes(action)) return;

  try {
    // console.log('[sandbox] start pptx generation');
    // 1) インスタンス生成
    const pptx = new PptxGenJS();

    // 2) カスタムスライドサイズ (16:9)
    pptx.defineLayout({ name: 'CUSTOM_LAYOUT', width: LAYOUT_WIDTH, height: LAYOUT_HEIGHT });
    pptx.layout = 'CUSTOM_LAYOUT';

    // 3) スライドマスター定義
    const logoPath = chrome.runtime && chrome.runtime.getURL ? chrome.runtime.getURL('logo.png') : 'logo.png';
    defineMasterSlide(pptx, logoPath);

    // 4) addSlide をラップしてマスタースライドを自動適用
    const originalAddSlide = pptx.addSlide.bind(pptx);
    pptx.addSlide = (options = {}) => {
      if (typeof options === 'string') {
        options = { masterName: options };
      } else if (typeof options === 'object' && !options.masterName) {
        options.masterName = MASTER_SLIDE_TITLE;
      }
      return originalAddSlide(options);
    };


    // addChart をラップしてデバッグ情報を記録
    try {
      const tmp = pptx.addSlide();
      const slideProto = Object.getPrototypeOf(tmp);
      const originalAddChart = slideProto.addChart;
      // グラフ追加時にエラーを検出するためのラッパー
      slideProto.addChart = function(chartType, data, opts) {
        try {
          const res = originalAddChart.call(this, chartType, data, opts);
          return res;
        } catch (e) {
          console.error('[sandbox][addChart] error', e);
          throw e;
        }
      };
      if (Array.isArray(pptx._slides)) pptx._slides.pop();
    } catch (err) {
      console.warn('[sandbox] failed to wrap addChart', err);
    }

    // 5) ユーザーコード実行
    // console.log('[sandbox] executing user code');
    if (action === 'generate-multi') {
      const arr = Array.isArray(codes) ? codes : [];
      for (const snip of arr) {
        try {
          const sanitized = sanitizeSnippet(snip);
          const processed = await replaceGetUrls(sanitized);
          if (!isSafeCode(processed)) {
            throw new Error('Unsafe code detected');
          }
            await runPptxFromSnippet(processed, { pptx });
        } catch (err) {
          console.error('[sandbox] snippet error', err);
          try {
            const errSlide = pptx.addSlide();
            errSlide.addText('エラーが起きてこのスライドの作成に失敗しました。', {
              x: 0.5,
              y: 0.5,
              w: LAYOUT_WIDTH - 1,
              h: 1,
              fontFace: DEFAULT_FONT_FACE,
              fontSize: 18,
              color: 'FF0000',
              bold: true,
            });
            errSlide.addText(err.message || String(err), {
              x: 0.5,
              y: 1.5,
              w: LAYOUT_WIDTH - 1,
              h: LAYOUT_HEIGHT - 2,
              fontFace: DEFAULT_FONT_FACE,
              fontSize: 14,
              color: '000000',
            });
          } catch (e2) {
            console.error('[sandbox] failed to add error slide', e2);
          }
        }
        normalizeColors(pptx);
      }
    } else {
      const processed = await replaceGetUrls(code);
      if (!isSafeCode(processed)) {
        throw new Error('Unsafe code detected');
      }
        await runPptxFromSnippet(processed, { pptx });
        normalizeColors(pptx);
    }

    // --- append additional slides ---
    try {
      const fontNote = pptx.addSlide();
      try {
        const imgRes = await fetch(chrome.runtime.getURL('images/IMG_2072.jpeg'));
        if (!imgRes.ok) {
          throw new Error(`status ${imgRes.status}`);
        }
        const imgBlob = await imgRes.blob();
        const imgData = await new Promise((resolve) => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result);
          reader.onerror = () => resolve(null);
          reader.readAsDataURL(imgBlob);
        });
        if (imgData) {
          fontNote.addImage({
            data: imgData,
            x: 0.45934055118110234,
            y: 1.6579800962379703,
            w: 12.238332239720036,
            h: 5.35167760279965,
          });
        }
      } catch (err) {
        console.warn('[sandbox] failed to load font note image', err);
      }
      fontNote.addText('注意点：フォントの埋め込み', {
        x: 0.36632108486439197,
        y: 0.3543307086614173,
        w: 12.598425196850394,
        h: 0.5901870078740158,
        fontFace: DEFAULT_FONT_FACE,
        fontSize: 22,
        bold: true,
      });
      fontNote.addText('BoxのプレビューやPDFではフォントがRakuten Sans JPから別のフォントに置き換わる。それを避けたい場合は下記の埋め込み設定をしてから保存してください。', {
        x: 0.3663188976377953,
        y: 0.923844050743657,
        w: 12.598426290463692,
        h: 0.5902777777777778,
        fontFace: DEFAULT_FONT_FACE,
        fontSize: 14,
      });

      const fin = pptx.addSlide();
      fin.addText([
        { text: '何か不具合や要望ががありましたら' },
        { text: 'こちら', options: { hyperlink: { url: INQUIRY_URL } } },
        { text: 'からご連絡ください。\n' },
        { text: 'また、マニュアルは' },
        { text: 'こちら', options: { hyperlink: { url: MANUAL_URL } } },
        { text: 'です。' }
      ], {
        x: 1,
        y: LAYOUT_HEIGHT - 1,
        w: LAYOUT_WIDTH - 2,
        fontSize: 14,
        fontFace: DEFAULT_FONT_FACE,
        align: 'right',
      });
    } catch (e) {
      console.warn('[sandbox] failed to add final slide', e);
    }

    // 6) Blob 生成 (compression: 'STORE' で圧縮なし)
    // console.log('[sandbox] writing pptx to blob');
    const blob = await writeWithFallback(pptx, options);

    if (action === 'generate-single') {
      safePostMessage({ action: 'single-generated', blob });
      return;
    }
    if (action === 'generate-multi') {
      safePostMessage({ action: 'multi-generated', blob });
      return;
    }

  // 7) Blob → DataURL
  const fr = new FileReader();
  fr.onprogress = (ev) => {
    if (ev.lengthComputable) {
      // console.log(`[sandbox] reading blob: ${((ev.loaded / ev.total) * 100).toFixed(0)}%`);
      safePostMessage({ action: 'progress', loaded: ev.loaded, total: ev.total });
    }
  };
  fr.onload = () => {
    // console.log('[sandbox] file reader done');
    safePostMessage({ action: 'done', dataURL: fr.result, fileName });
  };
  fr.onerror = (err) => {
    console.error('[sandbox] file reader error', err);
    safePostMessage({ action: 'error', message: err.message || 'FileReader error' });
  };
  fr.onabort = () => {
    console.error('[sandbox] file reader aborted');
    safePostMessage({ action: 'error', message: 'FileReader aborted' });
  };
  fr.readAsDataURL(blob);
  } catch (err) {
    console.error('[sandbox] generation error', err);
    safePostMessage({ action: 'error', message: err.message });
  }
});
