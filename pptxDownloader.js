/**
 * ファイル名: pptxDownloader.js
 * 説明:
 *   プレビューパネルからPPTXファイルをダウンロードするための制御スクリプト。
 *   ユーザーが生成した pptxgenjs コードをサンドボックス化した iframe に送り、
 *   進捗表示・キャンセル処理・ファイル保存をまとめて行う。
 *
 * 主な処理の流れ:
 *   1. ボタン操作を受けてダウンロード処理を開始し、アイコンパスを DataURL に変換する replaceGetUrls を適用。
 *   2. hidden iframe を動的に生成してコードを postMessage で送信し、生成中の進捗メッセージを受信して UI に反映。
 *   3. 完了メッセージでは DataURL を a 要素に設定してダウンロードを実行、エラー時はユーザーに警告を表示。
 *   4. キャンセルや完了後はイベントリスナーやボタン状態をリセットして後始末を行う。
 */
(() => {
  // プレビューパネルの準備を行う初期化処理
  function init() {
    const app = window.previewApp;
    if (!app) return;
    if (typeof chrome === 'undefined' || !chrome.runtime || !chrome.runtime.getURL) {
      window.chrome = {
        runtime: { getURL: (path) => `${location.origin}/${path.replace(/^\//, '')}` },
      };
    }
    const FALLBACK_ICON_PATH = 'icon/solid/square.svg';
    const iconCache = new Map();

  // 指定したアイコンファイルを読み込み DataURL で返す
  async function loadIcon(path, colorHex, toURL) {
    const cacheKey = `${path}:${colorHex || ''}`;
    if (iconCache.has(cacheKey)) {
      return iconCache.get(cacheKey);
    }
    const url = toURL(path);
    const res = await fetch(url);
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

  // アイコン取得に失敗した場合はフォールバックを適用
  async function processIcon(path, colorHex, toURL) {
    const data = await loadIcon(path, colorHex, toURL).catch(() => null);
    if (data) return data;
    return (await loadIcon(FALLBACK_ICON_PATH, colorHex, toURL).catch(() => null))
      || `'${toURL(FALLBACK_ICON_PATH)}'`;
  }

  // コード中の chrome.runtime.getURL 呼び出しを実データURLに差し替える
  async function replaceGetUrls(code) {
    if (typeof chrome === 'undefined' || !chrome.runtime || !chrome.runtime.getURL) {
      return code;
    }

    const safeGetURL = (() => {
      let getURL = chrome?.runtime?.getURL;
      if (getURL) {
        try {
          getURL('');
        } catch (e) {
          getURL = null;
        }
      }
      return (p) => (getURL ? getURL(p) : null);
    })();

    const regex = /chrome\.runtime\.getURL\(\s*(['"])([^'"]+)\1\s*\)/g;
    const matches = [...code.matchAll(regex)];
    const promises = matches.map(async (m) => {
      const original = m[2];
      const [path, query] = original.split('?');
      const params = new URLSearchParams(query || '');
      const color = params.get('color');
      const colorHex = color ? (color.startsWith('#') ? color : `#${color}`) : null;
      const url = safeGetURL(path);
      if (!url) {
        return `'${original}'`;
      }
      return await processIcon(path, colorHex, safeGetURL);
    });

    const results = await Promise.all(promises);
    let replaced = code;
    matches.forEach((m, idx) => {
      replaced = replaced.replace(m[0], results[idx]);
    });
    return replaced;
  }

  // エクスポート失敗時のメッセージをローカライズして取得
  function getExportFailedMessage() {
    try {
      const lang = (typeof localStorage !== 'undefined' && localStorage.getItem('lang')) || 'ja';
      return (
        window.messages?.[lang]?.exportFailed ||
        'エラーが発生しエクスポートがうまくいきませんでした。修正指示をAIに渡してください'
      );
    } catch {
      return 'エラーが発生しエクスポートがうまくいきませんでした。修正指示をAIに渡してください';
    }
  }

  // hidden iframe を用いて PPTX を生成しダウンロードする
  function downloadPptx() {
    if (!app.scrapedCode) return;
    const btn = document.getElementById('download-btn');
    if (btn) {
      btn.disabled = true;
      btn.classList.remove('active');
    }
    let fileName = 'presentation.pptx';
    try {
      const m = app.scrapedCode.match(/addText\(\s*(["'`])([\s\S]*?)\1/);
      if (m) {
        const text = m[2].replace(/\\n/g, ' ').trim();
        const sanitized = text.replace(/[\\/:*?"<>|]/g, '_').slice(0, 20);
        const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        fileName = `${sanitized}_${date}.pptx`;
      } else {
      }
    } catch (e) {
      console.error('[downloader] failed to extract addText text', e);
    }
    app.showProgress('PowerPointエクスポート中...', 120000);

    let canceled = false;
    app.cancelDownload = () => {
      if (!canceled) {
        canceled = true;
        cleanup();
      }
    };

    const send = async () => {
      if (canceled) return;
      const code = await replaceGetUrls(app.scrapedCode);
      if (canceled) return;
      postToIframe(code);
    };

    if (!app.downloadIframe) {
      app.downloadIframe = document.createElement('iframe');
      app.downloadIframe.id = app.BOX_ID;
      app.downloadIframe.style.display = 'none';
      app.downloadIframe.setAttribute('sandbox', 'allow-scripts');
      const src = chrome && chrome.runtime && chrome.runtime.getURL
        ? chrome.runtime.getURL('sandbox/pptx-sandbox.html')
        : 'sandbox/pptx-sandbox.html';
      app.downloadIframe.src = src;
      app.downloadIframe.onload = () => {
        app.iframeReady = true;
        send();
      };
      document.body.appendChild(app.downloadIframe);
    } else if (!app.iframeReady) {
      app.downloadIframe.onload = () => {
        app.iframeReady = true;
        send();
      };
    } else {
      send();
    }

    // iframe にコードを送り生成処理を開始
    function postToIframe(code) {
      if (canceled) return;
      window.removeEventListener('message', handleDownload);
      window.addEventListener('message', handleDownload);
      const iframeOrigin = new URL(app.downloadIframe.src, window.location.href).origin;
      const sandboxed =
        app.downloadIframe.sandbox &&
        !app.downloadIframe.sandbox.contains('allow-same-origin');
      const target = sandboxed ? '*' : iframeOrigin;
      app.downloadIframe.contentWindow.postMessage(
        { action: 'generate', code: code, fileName },
        target
      );
    }

    // sandbox からのメッセージを受信して進捗や完了を処理
    function handleDownload(e) {
      const iframe = app.downloadIframe;
      if (!iframe || e.source !== iframe.contentWindow) return;
      const iframeOrigin = new URL(iframe.src, window.location.href).origin;
      const sandboxed = iframe.sandbox && !iframe.sandbox.contains('allow-same-origin');
      if (!sandboxed && e.origin !== iframeOrigin) return;
      if (sandboxed && e.origin !== 'null') return;
      const m = e.data || {};
      if (canceled) return;
      if (m.action === 'progress') {
        const ratio = m.total ? Math.round((m.loaded / m.total) * 100) : 0;
        app.updateProgress(Math.max(app.progressValue, ratio));
      }
      if (m.action === 'done') {
        app.stopFakeProgress();
        app.updateProgress(100);
        saveDataURL(m.dataURL, m.fileName);
        cleanup();
        if (typeof app.detectPptxAndRender === 'function') {
          app.detectPptxAndRender();
        }
      }
      if (m.action === 'error') {
        console.error('[downloader] error: ' + m.message);
        app.stopFakeProgress();
        if (typeof app.updateProgressMessage === 'function') {
          app.updateProgressMessage('exportFailed');
        }
        alert(getExportFailedMessage());
        cleanup();
        if (typeof app.insertErrorPrompt === 'function') {
          setTimeout(() => app.insertErrorPrompt(m.message), 700);
        }
      }
    }

    // DOM や進捗表示を片付ける
    function cleanup() {
      app.stopFakeProgress();
      window.removeEventListener('message', handleDownload);
      if (btn) {
        btn.disabled = false;
        btn.classList.add('active');
      }
      app.hideProgress();
      app.cancelDownload = null;
    }

    // DataURL をリンク経由で保存
    function saveDataURL(url, name) {
      const a = document.createElement('a');
      a.href = url;
      a.download = name;
      document.body.appendChild(a);
      a.click();
      a.remove();
    }
  }
  app.downloadPptx = downloadPptx;

  // Expose utilities for PPTX preview functionality
  window.pptxDownloaderUtils = {
    replaceGetUrls: replaceGetUrls
  };
}
if (window.previewApp) {
  init();
} else {
  document.addEventListener('previewAppReady', init, { once: true });
}
})();
