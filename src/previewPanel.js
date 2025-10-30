/**
 * ファイル名: src/previewPanel.js
 * 説明:
 *   プレビューパネル全体を制御するメインスクリプト。
 *   RakutenAI から受け取った HTML や pptxgenjs コードを表示・変換し、
 *   ユーザー操作に応じてスライドのブラッシュアップやエクスポートを行う。
 *
 * 主な役割:
 *   - 許可されたオリジンのみで動作するようチェックし、safeGetURL/resolvePreviewAssets で拡張内リソースに差し替え。
 *   - 多言語対応用の messages を参照し、applyTranslations で UI テキストを切り替える。
 *   - openTab などのブラウザ機能呼び出し、プログレス表示、ボタン・ドロップアップメニューなどの UI 操作をまとめて管理。
 *   - secureStorage による一時的な暗号化保存や、AI との postMessage 通信を仲介しながら pptxDownloader などへ処理を委譲。
 *   - シングルスライド/マルチスライド両モードでプレビューを行い、必要に応じて sandbox を介して PPTX を生成する。
 */
(async () => {
  if (window.previewApp) return;
  if (typeof chrome === 'undefined' || !chrome.runtime || !chrome.runtime.getURL) {
    window.chrome = {
      runtime: { getURL: (path) => `${location.origin}/${path.replace(/^\//, '')}` },
    };
  }
  const {
    initAIClient,
    startChat,
    startPromptList,
    autoSendPrompt,
    autoSendPreviewHtml,
    sendSlidesFromJson,
    sendSingleMessage,
    ensureResearchOff,
    ensureResearchOn,
    ensureWebSearchOff,
    monitorDeepResearchLoading,
    logStep,
    injector,
    findElement,
    sleep,
    waitForAnswerComplete,
  } = await import(chrome.runtime.getURL('src/aiClient.js'));

  const {
    secureStorage,
    loadRegStatus,
    saveRegStatus,
    loadPanelState,
    loadAnnouncementState,
    saveAnnouncementState,
    savePanelState,
    handOff,
    loadPayload,
  } = await import(chrome.runtime.getURL('src/storage.js'));
  const payload = await loadPayload();

  const PANEL_ELEMENT_ID = 'custom-preview-panel';

  // お知らせ内容の更新手順は docs/お知らせ更新ガイド.md を参照してください。
  const CURRENT_ANNOUNCEMENT = {
    id: '2025093001',
    date: '2025.09.30',
    version: 'Version 3.8.1',
    subtitle: 'CSV流し込み機能でスライド量産',
    body: [
      'CSVファイルを読み込ませて、デザインを保ったままコンテンツだけ一気に差し替えることが可能です。',
      '同じレイアウトで都道府県ごとのスライドを作成するなどの際に活用できます。',
      '<a class="announcement-text-link" href="https://rak.box.com/s/qzew0b26hmsho8742czuojkndl5uookf" target="_blank" rel="noreferrer noopener">詳しくはこちら</a>',
    ],
    actionLabel: '閉じる',
    imageSrc: 'images/Explanation.svg',
    defaultLang: 'ja',
    translations: {
      ja: {
        actionLabel: '閉じる',
      },
      en: {
        date: 'September 30, 2025',
        version: 'Version 3.8.1',
        subtitle: 'Scale slide production with CSV imports',
        body: [
          'Load a CSV to swap content in bulk while keeping your slide design intact.',
          'Ideal for producing prefecture-specific decks or any repeated layout.',
          '<a class="announcement-text-link" href="https://rak.box.com/s/qzew0b26hmsho8742czuojkndl5uookf" target="_blank" rel="noreferrer noopener">Learn more</a>',
        ],
        actionLabel: 'OK',
      },
    },
  };

  function getAnnouncementCopy(lang = currentLang) {
    if (!CURRENT_ANNOUNCEMENT || !CURRENT_ANNOUNCEMENT.id) return null;
    const base = {
      date: CURRENT_ANNOUNCEMENT.date || '',
      version: CURRENT_ANNOUNCEMENT.version || '',
      subtitle: CURRENT_ANNOUNCEMENT.subtitle || '',
      body: Array.isArray(CURRENT_ANNOUNCEMENT.body)
        ? CURRENT_ANNOUNCEMENT.body
        : CURRENT_ANNOUNCEMENT.body
          ? [String(CURRENT_ANNOUNCEMENT.body)]
          : [],
      actionLabel: CURRENT_ANNOUNCEMENT.actionLabel || '',
    };

    const translations = CURRENT_ANNOUNCEMENT.translations || {};
    const normalized = (lang || '').toLowerCase();
    const languageCandidates = [];
    if (normalized) languageCandidates.push(normalized);
    if (normalized.includes('-')) {
      const primary = normalized.split('-')[0];
      if (!languageCandidates.includes(primary)) languageCandidates.push(primary);
    }
    const defaultLang = CURRENT_ANNOUNCEMENT.defaultLang;
    if (defaultLang && !languageCandidates.includes(defaultLang)) {
      languageCandidates.push(defaultLang);
    }
    Object.keys(translations).forEach((code) => {
      if (!languageCandidates.includes(code)) languageCandidates.push(code);
    });

    let localized = null;
    for (const code of languageCandidates) {
      if (translations[code]) {
        localized = translations[code];
        break;
      }
    }

    if (!localized) return base;

    const body = Array.isArray(localized.body)
      ? localized.body
      : localized.body
        ? [String(localized.body)]
        : base.body;

    return {
      ...base,
      ...localized,
      body,
    };
  }

  function updateAnnouncementLocale(root, lang = currentLang) {
    if (!root || !CURRENT_ANNOUNCEMENT || !CURRENT_ANNOUNCEMENT.id) return;
    const modal = root.querySelector('#announcement-modal');
    if (!modal) return;

    const copy = getAnnouncementCopy(lang);
    if (!copy) return;

    const dateEl = modal.querySelector('[data-announcement-slot="date"]');
    const titleEl = modal.querySelector('[data-announcement-slot="version"]');
    const subtitleEl = modal.querySelector('[data-announcement-slot="subtitle"]');
    const bodyEl = modal.querySelector('[data-announcement-body]');
    const actionEl = modal.querySelector('#announcement-confirm');

    if (dateEl) dateEl.textContent = copy.date || '';
    if (titleEl) titleEl.innerHTML = copy.version || '';
    if (subtitleEl) subtitleEl.innerHTML = copy.subtitle || '';
    if (bodyEl) {
      bodyEl.innerHTML = '';
      (copy.body || []).forEach((line) => {
        const paragraph = document.createElement('p');
        paragraph.className = 'announcement-text';
        paragraph.innerHTML = line;
        bodyEl.appendChild(paragraph);
      });
    }
    if (actionEl) actionEl.innerHTML = copy.actionLabel || '';
  }

  const ALLOWED_ORIGINS = [
    'https://r-ai.tsd.public.rakuten-it.com',
  ];
  const BASE_ORIGIN =
    ALLOWED_ORIGINS.find(o => location.href.startsWith(o)) ||
    ALLOWED_ORIGINS.find(o => document.referrer.startsWith(o)) ||
    ALLOWED_ORIGINS[0];
  const URL_MAP = {
    'https://r-ai.tsd.public.rakuten-it.com': {
      CHAT_URL: 'https://r10.to/h5ATT8',
      START_CHAT_URL: 'https://r10.to/hkXhgq',
    },
  };
  const urlConfig = URL_MAP[BASE_ORIGIN] || URL_MAP[ALLOWED_ORIGINS[0]];

  const safeGetURL = (p) => {
    try {
      return chrome?.runtime?.getURL ? chrome.runtime.getURL(p) : p;
    } catch {
      return p;
    }
  };

  const resolvePreviewAssets = (html) => {
    if (!html) return html;

    const iconToDataURL = (origPath) => {
      const url = new URL(origPath, location.origin);
      const color = url.searchParams.get('color');
      const basePath = url.pathname.replace(/^\//, '');
      const fullPath = safeGetURL(basePath);
      try {
        const xhr = new XMLHttpRequest();
        xhr.open('GET', fullPath, false);
        xhr.send();
        if (xhr.status < 200 || xhr.status >= 300) throw new Error('not found');
        let svg = xhr.responseText;
        if (color) {
          const hex = `#${color}`;
          svg = svg
            .replace(/<svg([^>]*)>/, (m, attrs) => `<svg${attrs} fill="${hex}" stroke="${hex}" style="color:${hex}">`)
            .replace(/fill="[^"]*"/g, `fill="${hex}"`)
            .replace(/stroke="[^"]*"/g, `stroke="${hex}"`);
        }
        const encoded = btoa(unescape(encodeURIComponent(svg)));
        return `data:image/svg+xml;base64,${encoded}`;
      } catch {
        const fallbackColor = color ? `#${color}` : '#ccc';
        const fallbackSvg = `<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><rect width='100' height='100' fill='${fallbackColor}'/></svg>`;
        const encoded = btoa(unescape(encodeURIComponent(fallbackSvg)));
        return `data:image/svg+xml;base64,${encoded}`;
      }
    };

    let resolved = html
      .replace(/\$\{safeGetURL\(['"]([^'"]+)['"]\)\}/g, (_, p) => safeGetURL(p))
      .replace(/(src|href)=["']([^"']*icon\/[^"']+\.svg(?:\?[^"']*)?)["']/g, (m, attr, path) => `${attr}="${iconToDataURL(path)}"`)
      .replace(/https:\/\/cdn\.jsdelivr\.net\/npm\/tailwindcss[^"']*/g, safeGetURL('other/tailwind.min.css'))
      .replace(/https:\/\/cdn\.jsdelivr\.net\/npm\/chart\.js[^"']*/g, safeGetURL('other/chart.js'));

    if (!/tailwind(?:\.min)?\.css/.test(resolved)) {
      resolved = resolved.replace(
        /<head[^>]*>/i,
        (m) => `${m}<link rel="stylesheet" href="${safeGetURL('other/tailwind.min.css')}">`
      );
    }

    if (!/chart(?:\.umd)?\.js/.test(resolved) && /<canvas/i.test(resolved)) {
      resolved = resolved.replace(
        /<head[^>]*>/i,
        (m) => `${m}<script src="${safeGetURL('other/chart.js')}"></script>`
      );
    }

    return resolved;
  };

  // Ensure panel-specific styles (including drop-up menu) are loaded
  const cssLink = document.createElement('link');
  cssLink.rel = 'stylesheet';
  cssLink.href = safeGetURL('previewPanel.css');
  document.head.appendChild(cssLink);

  const messages = window.messages || {};
  let currentLang = (typeof localStorage !== 'undefined' && localStorage.getItem('lang')) || 'ja';

  const t = (key) => messages[currentLang]?.[key] || key;

  // 現在の言語設定に基づき、UI 要素に翻訳済みテキストを挿入する
  function applyTranslations(root = document) {
    root.querySelectorAll('[data-i18n]').forEach((el) => {
      const key = el.dataset.i18n;
      if (key) el.textContent = t(key);
    });
    root.querySelectorAll('[data-i18n-placeholder]').forEach((el) => {
      const key = el.dataset.i18nPlaceholder;
      if (key) el.placeholder = t(key);
    });
    const btn = document.getElementById('lang-btn');
    if (btn) {
      // 言語選択ボタン自体も現在の言語名に置き換える
      btn.innerHTML =
        '<img src="' +
        safeGetURL(app.LANGUAGE_ICON_SRC) +
        '" alt="">' +
        t('languageName') +
        ' ▴';
    }
  }

  // 選択された言語を保存し UI を再描画する
  function setLanguage(lang) {
    currentLang = lang;
    try { localStorage.setItem('lang', lang); } catch {}
    applyTranslations();
    const panel = document.getElementById(PANEL_ELEMENT_ID);
    if (panel) updateAnnouncementLocale(panel, lang);
  }

  // ホワイトリストに基づき URL を新しいタブで開くヘルパー
  function openTab(url, active = false) {
    const WHITELIST = [
      ...ALLOWED_ORIGINS.map((o) => `${o}/`),
      'https://forms.office.com/',
      'https://chromewebstore.google.com/',
      'https://r10.to/',
    ];
    if (!WHITELIST.some((p) => url.startsWith(p))) {
      console.warn('Blocked openTab to non-whitelisted URL:', url);
      return;
    }
    try {
      if (chrome?.runtime?.sendMessage) {
        // 拡張機能のバックグラウンドに委譲してタブを開く
        chrome.runtime.sendMessage({ action: 'open-tab', url, active });
        return;
      }
    } catch (e) {
      console.error('openTab error', e);
    }
    const feats = active ? '' : 'noopener';
    window.open(url, '_blank', feats);
  }

  const PRIMARY_COLOR = '#1e88e5';
const PRIMARY_HOVER = '#1565c0';
const CONVERT_PREFIX = `コードをブラッシュアップをして.デザインを美しく見やすくしてください。細部までこだわって。グラフがある場合は特に見やすく美しくしてください。Titleと見出しは左揃えにしてください。

Font Awesome アイコンの <i> タグを検出したら、拡張機能の icon フォルダ内にある対応する SVG ファイルをchrome.runtime.getURL() で参照する slide.addImage() 呼び出しに変換するJavaScriptコードを出力してください。  ◎ 変換ルール 1. <i class="fas fa-ICON-NAME ..."></i>    → slide.addImage({        path: chrome.runtime.getURL(\`icon/solid/ICON-NAME.svg\`),        x: X, y: Y, w: W, h: H      }); 2. <i class="far fa-ICON-NAME ..."></i>    → slide.addImage({        path: chrome.runtime.getURL(\`icon/regular/ICON-NAME.svg\`),        x: X, y: Y, w: W, h: H      }); 3. <i class="fab fa-ICON-NAME ..."></i>    → slide.addImage({        path: chrome.runtime.getURL(\`icon/brands/ICON-NAME.svg\`),        x: X, y: Y, w: W, h: H      });  ◎ 要件 - HTML 中の複数パターン（順序の異なるクラス混在を含む）にも対応 - X,Y,W,H は適宜スライドのレイアウトに合わせて変数化 - 色指定がある<i>タグは、path の末尾に ?color=HEX を追加する  アイコンのカラーの指定方法は“?color=HEX”です。  例: slide.addImage({   path: chrome.runtime.getURL('icon/solid/check-circle.svg?color=ff0000'),   x: X1, y: Y1, w: W1, h: H1 }); slide.addImage({   path: chrome.runtime.getURL('icon/regular/smile.svg?color=00ff00'),   x: X2, y: Y2, w: W2, h: H2 }); slide.addImage({   path: chrome.runtime.getURL('icon/brands/github.svg'), // デフォルト色（黒）   x: X3, y: Y3, w: W3, h: H3 });

絶対にパワポを開いたときに壊れていないように確認してPptxgenjsのコードを適切に書いて。`;
const MULTI_SLIDE_PREFIX = `あなたは、与えられたトピックに基づいてプレゼンテーションスライドを自動生成するAIです。ユーザーが要件（トピック）を入力すると、以下の手順を実行してください。



トピックからスライドの構成案（セクションやポイント）をまとめます。


各スライドの要点を整理し、スライド順に並べます。

最初の1枚目のスライドは後続の内容をまとめたエグゼクティブサマリーにしてください。


以下の形式で出力してください。
1. 人間が読みやすいテキストでの構造的に整理された構成案
2. 上記の後ろにJSON形式の出力

JSONでは slides 配列に、各スライド用のテキストプロンプトを収めます。プロンプトは詳細に書いてください。改行が必要な場合は\nを含めてください。

例:

{


  "slides": [


    {


      "title": "エグゼクティブサマリー",


      "prompt": "後続スライドの内容をまとめた全体概要を示すスライド"


    },


    {


      "title": "メインポイント1",


      "prompt": "ポイント1の詳細を説明するスライド"


    }


  ]


}
必ずユーザーのOKが出るまで構成案を提案し続けてください。構成案がOKなら「スライド作成開始」ボタンをクリックしてくださいね。ユーザーの入力は下記です。-----`;
const JSON_TO_HTML_PREFIX =
  '下記の内容を元にHTMLで16:9スライドを作成してください。コードブロックの形式で、フルのHTMLで書いてください。';

  const app = {
    PANEL_ID: PANEL_ELEMENT_ID,
    STYLE_ID: 'custom-preview-style',
    TOGGLE_ID: 'custom-preview-toggle',
    ACTIVE_CLS: 'custom-preview-active',
    PANEL_W: '50%',
    BOX_ID: 'pptx-sandbox-frame',
    CHAT_URL: urlConfig.CHAT_URL,
    START_CHAT_URL: urlConfig.START_CHAT_URL,
    CHAT_CONTAINER_ID: 'external-chat-container',
    PREVIEW_PARAM: 'autoPreview',
    HTML_PARAM: 'previewHtml',
    HTML_LIST_PARAM: 'htmlSlides',
    PROMPT_PARAM: 'prompt',
    PROMPT_LIST_PARAM: 'promptList',
    AUTO_DL_PARAM: 'autoDownload',
    DEEP_PARAM: 'deepResearch',
    SEARCH_PARAM: 'searchMode',
    SLIDE_PARAM: 'slidePrompt',
    WAIT_PARAM: 'wait',
    HOME_BTN_ID: 'home-btn',
    HOME_GRAY_SRC: 'images/home-gray.svg',
    HOME_WHITE_SRC: 'images/home-white.svg',
    CONVERT_ICON_SRC: 'icon/solid/file-powerpoint.svg',
    DOWNLOAD_ICON_SRC: 'icon/solid/download.svg',
    CLOSE_ICON_SRC: 'icon/regular/window-close.svg',
    LANGUAGE_ICON_SRC: 'icon/solid/globe.svg',
    INQUIRY_URL: 'https://forms.office.com/r/04iMPYrVmv',
    MANUAL_URL: 'https://rak.box.com/s/4dvyt5n5vbvq1lbdwo36n8hdqf527qg1',
    SHARE_URL: 'https://r10.to/hktakv',

    scrapedCode: '',
    lastHtml: '',
    lastPptx: '',
    observer: null,
    renderTimer: null,
    downloadIframe: null,
    iframeReady: false,
    progressValue: 0,
    downloading: false,
    fakeProgressTimer: null,
    chatFrame: null,
    prevDisabled: null,
    previewRegDone: false,
    pptxRegDone: false,
    regHideTimer: null,
    multiSlideMode: false,
    userSlideModeSet: false,
    htmlMonitorTimer: null,
    htmlSlides: [],
    slidePrompt: '',
    searchMode: "off",
    deepMonitorStarted: false,
    startText: '',
    isTyping: false,
    pendingRender: false,
    showHome: false,
    autoDlTimer: null,
    arrowTimer: null,
    creatingSlides: false,
    currentPreview: '',
    panelOpen: false,
    promptCollapsed: false,
    detectedHtml: false,
    announcementState: {},
    announcementDismissed: false,
    logStart: payload.logStart || 0,
    payload,

    // PPTX Preview related properties
    pptxViewer: null,
    pptxPreviewResult: null,
    pptxBlob: null,
  };

  // 変換ボタンにアイコンと翻訳テキストを設定する
  function updateConvertBtn(btn, key) {
    if (!btn) return;
    btn.innerHTML = '<img src="' + safeGetURL(app.CONVERT_ICON_SRC) + '" style="width:14px;height:14px;margin-right:4px;"><span data-i18n="' + key + '"></span>';
    applyTranslations(btn);
  }

  const urlParamMode = new URLSearchParams(location.search).get(app.SEARCH_PARAM);
  if (urlParamMode === 'websearch' || urlParamMode === 'web') {
    app.searchMode = 'web';
  } else if (urlParamMode === 'deepresearch' || urlParamMode === 'deep') {
    app.searchMode = 'deep';
  }

  // 検索モードをオフに戻し必要に応じて保存
  function resetSearchMode(persist = false) {
    app.searchMode = 'off';
    const sel = document.querySelector(`#${app.PANEL_ID} #search-mode-select`);
    if (sel) sel.value = 'off';
    if (persist) saveRegStatus(app);
  }

  function syncAnnouncementState(state) {
    if (state && state.id === CURRENT_ANNOUNCEMENT.id) {
      app.announcementState = state;
      app.announcementDismissed = !!state.dismissed;
    } else {
      app.announcementState = { id: CURRENT_ANNOUNCEMENT.id, dismissed: false };
      app.announcementDismissed = false;
    }
  }

  // ======================
  // PPTX Preview Functions
  // ======================

  let pptxPreviewLib = null;

  // Get pptx-preview library (already loaded via content_scripts in manifest)
  async function loadPptxPreview() {
    if (pptxPreviewLib) {
      console.log('[Preview] Using cached pptx-preview library');
      return pptxPreviewLib;
    }

    console.log('[Preview] Getting pptx-preview library from global scope...');

    // The library is loaded via content_scripts in manifest.json
    // so it should already be available in the global scope
    if (typeof pptxPreview !== 'undefined' && typeof pptxPreview.init === 'function') {
      pptxPreviewLib = pptxPreview;
      console.log('[Preview] pptx-preview library found successfully');
      return pptxPreviewLib;
    } else {
      console.error('[Preview] pptx-preview library not found in global scope');
      console.error('[Preview] typeof pptxPreview:', typeof pptxPreview);
      throw new Error('pptx-preview library is not available. Make sure lib/pptx-preview.iife.js is loaded.');
    }
  }

  // Generate PPTX Blob from code using API
  async function generatePptxBlob(code) {
    try {
      // Import API client module
      const { generatePptxBlobViaApi, getApiKey } = await import(chrome.runtime.getURL('src/apiClient.js'));

      // Check if API key is set
      const apiKey = await getApiKey();
      if (!apiKey) {
        throw new Error('APIキーが設定されていません。設定画面からAPIキーを設定してください。');
      }

      // Generate PPTX via API
      const blob = await generatePptxBlobViaApi(code, 'preview.pptx');
      return blob;
    } catch (error) {
      console.error('PPTX generation via API failed:', error);
      throw error;
    }
  }

  // Remove placeholder texts from preview
  function removePptxPlaceholders() {
    const view = document.getElementById('preview-content');
    if (!view) return;

    const container = view.querySelector('.preview-container');
    if (!container) return;

    // Hide placeholder text elements
    const textElements = container.querySelectorAll('svg text');
    const placeholderTexts = [
      '图表标题', '圖表標題', 'Chart Title',
      'Click to add title', 'Click to add text',
      '系列1', '系列2', '系列3', '系列4',
      'Series 1', 'Series 2', 'Series 3', 'Series 4',
      'Legend'
    ];

    textElements.forEach(textElement => {
      const textContent = textElement.textContent || '';
      if (placeholderTexts.includes(textContent.trim())) {
        textElement.remove();
      }
    });

    // Hide chart legends (凡例) - only hide legend elements, not chart grid/axes
    const svgElements = container.querySelectorAll('svg');
    svgElements.forEach(svg => {
      // Find all path elements that are legend markers
      const paths = svg.querySelectorAll('path');
      paths.forEach(path => {
        const d = path.getAttribute('d') || '';
        const transform = path.getAttribute('transform') || '';
        const stroke = path.getAttribute('stroke') || '';
        const strokeWidth = path.getAttribute('stroke-width') || '';

        // Legend marker patterns (avoid grid lines and axes):
        // 1. Short horizontal lines with transform: "M0 7L25 7" + translate (legend line marker)
        //    Grid lines don't have transform and are much longer
        // 2. Small circles with transform: "M18.1 7A5.6..." or "M1 0A1 1..." with translate
        // 3. Rounded rectangles for legend boxes: "M3.5 0L21.5 0A3.5 3.5..."

        // Exclude grid lines (no transform, stroke=#E0E6F1 or #6E7079, long paths)
        const isGridLine = !transform && (stroke === '#E0E6F1' || stroke === '#6E7079') && d.length > 30;

        if (!isGridLine) {
          const isLegendLine = transform.includes('translate') &&
                               d.match(/^M0 \d+L\d+ \d+$/) &&
                               d.length < 30 &&
                               strokeWidth === '2';

          const isLegendCircle = d.includes('A') &&
                                transform.includes('translate') &&
                                d.length < 100 &&
                                (d.includes('M1 0A1 1') || d.includes('M18'));

          const isLegendBox = d.includes('A') &&
                             d.includes('L') &&
                             d.length < 200 &&
                             (d.includes('3.5') || d.includes('10.5'));

          if (isLegendLine || isLegendCircle || isLegendBox) {
            path.style.display = 'none';
          }
        }
      });

      // Hide text elements that are legend labels only
      const texts = svg.querySelectorAll('text');
      texts.forEach(text => {
        const content = (text.textContent || '').trim();
        const transform = text.getAttribute('transform') || '';
        const x = parseFloat(text.getAttribute('x') || '0');

        // Legend text patterns - be very specific:
        // 1. Text with x="30" AND transform with translate(145 67.4) or similar legend positions
        // 2. Text content matches exact legend labels
        const isLegendLabel = (
          (x === 30 && transform.includes('translate') &&
           (transform.includes(' 67') || transform.includes(' 68'))) ||
          content === '売上高（億円）' ||
          content === '非常に満足' ||
          content === '満足' ||
          content === '普通' ||
          content === '不満'
        );

        if (isLegendLabel) {
          text.style.display = 'none';
        }
      });

      // Hide the bounding boxes for legend areas only
      svg.querySelectorAll('path').forEach(path => {
        const d = path.getAttribute('d') || '';
        const fill = path.getAttribute('fill') || '';
        const stroke = path.getAttribute('stroke') || '';

        // Very specific legend bounding boxes:
        // "M-5 -5l125 0l0 23.2l-125 0Z" with fill-opacity="0" or transparent
        if (d.startsWith('M-5 -5l') && d.includes('l-125 0Z') &&
            (fill.includes('0,0,0') || fill === 'transparent')) {
          path.style.display = 'none';
        }

        // Legend region paths: "M-1 0.4l115 0l0 13.2l-115 0Z" with fill="none"
        if (d.match(/^M-1 [\d.]+l\d+ 0l0 [\d.]+l-\d+ 0Z$/) && fill === 'none') {
          path.style.display = 'none';
        }
      });
    });

    // Hide chart titles (タイトル) - typically larger text at top of charts
    textElements.forEach(textElement => {
      const fontSize = parseFloat(window.getComputedStyle(textElement).fontSize || '0');
      const textContent = (textElement.textContent || '').trim();

      // Chart titles are typically:
      // 1. Larger font size (> 14px)
      // 2. Short text (less than 50 characters)
      // 3. Not empty
      if (fontSize > 14 && textContent.length > 0 && textContent.length < 50) {
        // Check if this text is positioned near the top of its SVG parent
        const svgParent = textElement.closest('svg');
        if (svgParent) {
          const y = parseFloat(textElement.getAttribute('y') || '0');
          const svgHeight = parseFloat(svgParent.getAttribute('height') || '0');

          // If positioned in top 30% of chart, likely a title
          if (y < svgHeight * 0.3) {
            textElement.style.display = 'none';
          }
        }
      }
    });
  }

  // Show PPTX preview in preview-container (same place as HTML preview)
  async function showPptxPreview() {
    console.log('[Preview] showPptxPreview called');

    if (!app.scrapedCode) {
      console.error('[Preview] No PPTX code to preview');
      alert(t('noTargetCode'));
      return;
    }

    // Get the panel first
    const panel = document.getElementById(app.PANEL_ID);
    if (!panel) {
      console.error('[Preview] Panel not found');
      return;
    }

    // Get the preview container (same as HTML preview)
    const view = panel.querySelector('#preview-content');
    if (!view) {
      console.error('[Preview] preview-content not found');
      return;
    }

    let container = view.querySelector('.preview-container');
    if (!container) {
      console.log('[Preview] .preview-container not found, creating it...');
      // Create the container if it doesn't exist
      container = document.createElement('div');
      container.className = 'preview-container';
      // Insert it as the first child of preview-content
      view.insertBefore(container, view.firstChild);
      console.log('[Preview] Container created');
    }

    console.log('[Preview] Container found:', container);

    // Set current preview type
    app.currentPreview = 'pptx';

    // Hide PPTX export overlay if it exists
    const pptxOverlay = view.querySelector('#pptx-detected-overlay');
    if (pptxOverlay) {
      pptxOverlay.style.display = 'none';
    }

    // Cleanup old viewer if it exists
    if (app.pptxViewer) {
      console.log('[Preview] Cleaning up old viewer...');
      try {
        if (typeof app.pptxViewer.destroy === 'function') {
          app.pptxViewer.destroy();
        }
      } catch (e) {
        console.warn('[Preview] Error destroying old viewer:', e);
      }
      app.pptxViewer = null;
      app.pptxPreviewResult = null;
    }

    // Clear container and show loading
    container.innerHTML = `
      <div class="preview-loading">
        <div class="spinner"></div>
        <p data-i18n="pptxPreviewLoading">プレビューを生成中...</p>
      </div>
    `;
    applyTranslations(container);

    try {
      console.log('[Preview] Step 1: Generating PPTX Blob via API...');
      // 1. Generate PPTX Blob
      const blob = await generatePptxBlob(app.scrapedCode);
      app.pptxBlob = blob;
      console.log('[Preview] PPTX Blob generated:', blob.size, 'bytes');

      console.log('[Preview] Step 2: Loading pptx-preview library...');
      // 2. Load pptx-preview library
      const lib = await loadPptxPreview();
      console.log('[Preview] Library loaded:', lib);

      if (!lib || !lib.init) {
        throw new Error('pptx-preview library does not have init method');
      }

      console.log('[Preview] Step 3: Converting to ArrayBuffer...');
      // 3. Convert to ArrayBuffer
      const arrayBuffer = await blob.arrayBuffer();
      console.log('[Preview] ArrayBuffer size:', arrayBuffer.byteLength);

      console.log('[Preview] Step 4: Creating preview wrapper and initializing...');
      // 4. Create wrapper div with 16:9 aspect ratio
      container.innerHTML = '';
      const wrapper = document.createElement('div');
      wrapper.className = 'pptx-preview-wrapper';
      container.appendChild(wrapper);

      // Calculate dimensions based on container width
      const containerWidth = container.clientWidth || 960;
      const previewHeight = containerWidth * (9 / 16);

      console.log('[Preview] Wrapper dimensions:', containerWidth, 'x', previewHeight);

      // Initialize viewer in the wrapper
      app.pptxViewer = lib.init(wrapper, {
        width: containerWidth,
        height: previewHeight
      });
      console.log('[Preview] Viewer initialized:', app.pptxViewer);

      console.log('[Preview] Step 5: Rendering preview...');
      // 5. Show preview
      app.pptxPreviewResult = await app.pptxViewer.preview(arrayBuffer);
      const totalSlides = app.pptxPreviewResult?.slides?.length ?? 0;
      console.log('[Preview] Total slides:', totalSlides);

      if (totalSlides === 0) {
        container.innerHTML = `
          <div class="preview-error">
            <p data-i18n="pptxPreviewError">プレビューの生成に失敗しました</p>
          </div>
        `;
        applyTranslations(container);
        return;
      }

      console.log('[Preview] Step 6: Removing placeholder texts...');
      // 6. Remove placeholder texts
      setTimeout(() => {
        removePptxPlaceholders();

        // Save preview HTML and show save button after placeholders are removed
        console.log('[Preview] Saving preview HTML...');
        let previewHtml = container.innerHTML;

        // Change background color from black to white for thumbnail display
        previewHtml = previewHtml.replace(/background:\s*rgb\(0,\s*0,\s*0\)/gi, 'background: rgb(255, 255, 255)');
        previewHtml = previewHtml.replace(/background:\s*#000000/gi, 'background: #FFFFFF');
        previewHtml = previewHtml.replace(/background:\s*black/gi, 'background: white');

        // Remove all margin values to prevent centering and spacing in thumbnail
        // Catches: margin: 0px auto, margin: 0px 10px, margin: 0px auto 10px, etc.
        previewHtml = previewHtml.replace(/margin:\s*[0-9.]+px\s+(auto|[0-9.]+px)(\s+[0-9.]+px)?(\s+[0-9.]+px)?/gi, 'margin: 0px');

        app.currentPreviewHtml = previewHtml;

        const savePreviewBtn = document.querySelector('#save-preview-btn');
        if (savePreviewBtn) {
          savePreviewBtn.style.display = 'inline-flex';
          savePreviewBtn.disabled = false;
          console.log('[Preview] Save button displayed');
        }
      }, 150);

      console.log('[Preview] Preview completed successfully!');

      // Update header buttons to change "プレビュー" to "プレビュー更新"
      setTimeout(() => {
        updateHeaderButtons();
      }, 200);

    } catch (error) {
      console.error('[Preview] Preview generation failed:', error);
      console.error('[Preview] Error stack:', error.stack);
      container.innerHTML = `
        <div class="preview-error">
          <p data-i18n="pptxPreviewError">プレビューの生成に失敗しました</p>
          <p style="font-size: 12px; color: #666;">${error.message}</p>
        </div>
      `;
      applyTranslations(container);
      app.currentPreview = ''; // Clear preview state on error
    }
  }

  function injectStyles() {
    const { PANEL_ID, STYLE_ID, TOGGLE_ID, PANEL_W, ACTIVE_CLS, CHAT_CONTAINER_ID } = app;
    if (document.getElementById(STYLE_ID)) return;
    const css = `
      #${PANEL_ID} {
        position: fixed;
        top: 0;
        right: 0;
        width: ${PANEL_W};
        height: 100vh;
        background: #f5f5f5;
        box-shadow: -2px 0 10px rgba(0,0,0,.2);
        z-index: 9999;
        display: flex;
        flex-direction: column;
        color-scheme: light;
      }
      body.${ACTIVE_CLS} {
        width: calc(100% - ${PANEL_W});
        margin: 0;
        overflow: hidden;
      }
      #${PANEL_ID} button {
        padding: 6px 12px;
        font-size: 14px;
        border: none;
        border-radius: 6px;
        background: ${PRIMARY_COLOR};
        color: #fff;
        cursor: pointer;
        box-shadow: 0 2px 4px rgba(0,0,0,0.15);
        transition: background 0.2s, box-shadow 0.2s, transform 0.1s, opacity 0.2s;
      }
      #${PANEL_ID} button:hover:not(:disabled) {
        background: ${PRIMARY_HOVER};
        box-shadow: 0 3px 6px rgba(0,0,0,0.2);
        transform: translateY(-1px);
      }
      #${PANEL_ID} button:active:not(:disabled) {
        transform: translateY(0);
      }
      #${PANEL_ID} button:disabled {
        opacity: 0.6;
        cursor: not-allowed;
      }
      #${TOGGLE_ID} {
        position: fixed;
        top: 50%;
        right: 0;
        transform: translateY(-50%);
        background: rgba(30, 136, 229, 0.6);
        color: #fff;
        border: none;
        border-top-left-radius: 5px;
        border-bottom-left-radius: 5px;
        padding: 4px;
        cursor: pointer;
        z-index: 9998;
        width: 32px;
        height: 80px;
        display: flex;
        align-items: center;
        justify-content: center;
      }
      #${TOGGLE_ID} img {
        width: 16px;
        transition: transform 0.3s;
      }
      #${TOGGLE_ID}.open {
        right: ${PANEL_W};
      }
      #${TOGGLE_ID}.open img {
        transform: rotate(180deg);
      }
      #${PANEL_ID} .preview-header {
        padding: 10px;
        border-bottom: 1px solid #ccc;
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: #eee;
      }
      #${PANEL_ID} .preview-header > div {
        display: flex;
        flex-wrap: wrap;
        gap: 4px 8px;
      }
      #${PANEL_ID} .preview-header button {
        margin: 0;
        height: 32px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
      }
      #${PANEL_ID} #start-all-btn.flashy {
        animation: glitter 1s infinite;
      }
      #${PANEL_ID} .preview-header button.active {
        background: ${PRIMARY_HOVER};
        color: #fff;
        font-weight: bold;
      }
      #${PANEL_ID} #home-btn {
        border: none;
        background: transparent;
        border-radius: 6px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        padding: 6px 12px;
        height: 32px;
        transition: background 0.2s;
      }
      #${PANEL_ID} #home-btn:hover {
        background: rgba(30, 136, 229, 0.15);
      }
      #${PANEL_ID} #home-btn.active {
        background: rgba(30, 136, 229, 0.3);
      }
      #${PANEL_ID} #home-btn img {
        width: 14px;
        height: 14px;
      }
      #${PANEL_ID} #convert-btn img,
      #${PANEL_ID} #download-btn img {
        filter: brightness(0) invert(1);
      }
      #${PANEL_ID} #start-chat-btn {
        transition: background 0.2s, color 0.2s, border-color 0.2s;
      }
      #${PANEL_ID} #start-chat-btn:active:not(:disabled) {
        transform: scale(0.97);
      }
      #${PANEL_ID} #start-chat-btn.active {
        font-weight: bold;
      }
      #${PANEL_ID} #start-chat-btn:disabled {
        opacity: 0.6;
        cursor: not-allowed;
      }
      #${PANEL_ID} .toggle-wrap {
        display: flex;
        align-items: center;
        margin-left: 8px;
      }
      #${PANEL_ID} .toggle-wrap .switch {
        position: relative;
        display: inline-block;
        width: 40px;
        height: 20px;
        margin-right: 4px;
      }
      #${PANEL_ID} .toggle-wrap .switch input {
        opacity: 0;
        width: 0;
        height: 0;
      }
      #${PANEL_ID} .toggle-wrap .slider {
        position: absolute;
        cursor: pointer;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background-color: #ccc;
        transition: .2s;
        border-radius: 20px;
      }
      #${PANEL_ID} .toggle-wrap .slider:before {
        position: absolute;
        content: '';
        height: 16px;
        width: 16px;
        left: 2px;
        bottom: 2px;
        background-color: white;
        transition: .2s;
        border-radius: 50%;
      }
      #${PANEL_ID} .toggle-wrap input:checked + .slider {
        background-color: ${PRIMARY_COLOR};
      }
      #${PANEL_ID} .toggle-wrap input:checked + .slider:before {
        transform: translateX(20px);
      }
      #${PANEL_ID} .toggle-wrap .toggle-label {
        font-size: 14px;
      }
      #${PANEL_ID} #search-mode-select {
        margin-left: 4px;
        padding: 4px 24px 4px 8px;
        font-size: 14px;
        border: 1px solid #ccc;
        border-radius: 4px;
        appearance: none;
        background-color: #fff;
        background-repeat: no-repeat;
        background-position: right 8px center;
        background-size: 10px 6px;
        background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' viewBox='0 0 10 6'%3E%3Cpath d='M1 1l4 4 4-4' stroke='%23999' stroke-width='2' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");
        cursor: pointer;
      }
      #${PANEL_ID} #slide-type-options {
        display: flex;
        gap: 10px;
        justify-content: center;
        margin: 5px 0 10px;
      }
      #${PANEL_ID} #slide-type-options .option {
        position: relative;
      }
      #${PANEL_ID} #slide-type-options img {
        width: 100px;
        height: auto;
        max-height: 70px;
        cursor: pointer;
        border: 2px solid transparent;
        border-radius: 4px;
        object-fit: contain;
        transition: transform 0.1s, border-color 0.2s;
      }
      #${PANEL_ID} #slide-type-options .label-overlay {
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 14px;
        font-weight: bold;
        color: #fff;
        background: rgba(0, 0, 0, 0.4);
        border-radius: 4px;
        opacity: 0;
        pointer-events: none;
        transition: opacity 0.2s;
      }
      #${PANEL_ID} #slide-type-options .option:hover .label-overlay {
        opacity: 1;
      }
      #${PANEL_ID} #slide-type-options img.selected {
        border-color: ${PRIMARY_COLOR};
      }
      #${PANEL_ID} #slide-type-options img:active {
        transform: scale(0.96);
      }
      #${PANEL_ID} #slide-mode-options {
        display: flex;
        flex-direction: column;
        align-items: stretch;
        margin: 5px 0;
        gap: 6px;
      }
      #${PANEL_ID} #slide-mode-options label {
        display: flex;
        align-items: center;
        justify-content: flex-start;
        gap: 8px;
        font-size: 14px;
        padding: 6px 8px;
        border: 1px solid #ccc;
        border-radius: 6px;
        cursor: pointer;
        transition: border-color 0.2s, box-shadow 0.2s;
      }
      #${PANEL_ID} #slide-mode-options label:hover {
        border-color: ${PRIMARY_COLOR};
        box-shadow: 0 0 0 2px ${PRIMARY_COLOR}33;
      }
      #${PANEL_ID} #slide-mode-options label .mode-title {
        font-weight: bold;
      }
      #${PANEL_ID} #slide-mode-options label .mode-desc {
        font-size: 12px;
        color: #666;
      }
      #${PANEL_ID} #slide-mode-options label .mode-text {
        display: flex;
        flex-direction: row;
        align-items: center;
        gap: 4px;
        text-align: left;
      }
      #${PANEL_ID} #slide-mode-options label .beta-tag {
        margin-left: auto;
        font-size: 10px;
        color: #fff;
        background: ${PRIMARY_COLOR};
        padding: 0 6px;
        border-radius: 3px;
      }
      #${PANEL_ID} .selected-tag {
        font-size: 10px;
        color: ${PRIMARY_COLOR};
        margin-left: 6px;
      }
      #${PANEL_ID} #slide-mode-options input {
        accent-color: ${PRIMARY_COLOR};
        transform: scale(1.2);
        margin-right: 6px;
      }
      #${PANEL_ID} #preview-content {
        flex: 1;
        min-height: 0;
        position: relative;
        background: #fff;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: flex-start;
      }
      #${PANEL_ID} #progress-indicator {
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(255, 255, 255, 0.8);
        display: none;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        z-index: 1000;
      }
      #${PANEL_ID} #wait-overlay {
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(255, 255, 255, 0.8);
        display: none;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        z-index: 3;
        font-size: 14px;
      }
      #${PANEL_ID} #arrow-overlay {
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(255, 255, 255, 0.3);
        display: none;
        z-index: 1;
      }
      #${PANEL_ID} #wait-overlay .wait-spinner {
        border: 4px solid #f3f3f3;
        border-top: 4px solid ${PRIMARY_COLOR};
        border-radius: 50%;
        width: 40px;
        height: 40px;
        animation: spin 1s linear infinite;
        margin-bottom: 8px;
      }
      #${PANEL_ID} #pptx-detected-overlay {
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(255, 255, 255, 0.9);
        display: none;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        z-index: 1;
        text-align: center;
        padding: 16px;
        box-sizing: border-box;
      }
      #${PANEL_ID} #pptx-detected-overlay .pptx-overlay-actions {
        margin-top: 10px;
        display: flex;
        flex-wrap: wrap;
        justify-content: center;
        gap: 8px;
      }
      #${PANEL_ID} #pptx-detected-overlay button {
        margin: 0;
      }
      #${PANEL_ID} #announcement-modal {
        position: absolute;
        inset: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 24px;
        background: rgba(15, 23, 42, 0.32);
        box-sizing: border-box;
        z-index: 10005;
        opacity: 0;
        pointer-events: none;
        transition: opacity 0.25s ease;
      }
      #${PANEL_ID} #announcement-modal.show {
        opacity: 1;
        pointer-events: auto;
      }
      #${PANEL_ID} #announcement-modal[aria-hidden="true"] {
        display: flex;
      }
      #${PANEL_ID} .announcement-modal-dialog {
        background: #fff;
        border-radius: 18px;
        box-shadow: 0 32px 64px rgba(15, 23, 42, 0.18);
        width: min(720px, 100%);
        padding: 32px;
        display: flex;
        flex-direction: column;
        gap: 28px;
        position: relative;
      }
      #${PANEL_ID} .announcement-modal-body {
        display: flex;
        gap: 28px;
      }
      #${PANEL_ID} .announcement-image {
        flex: 0 0 220px;
        background: #f8fafc;
        border-radius: 12px;
        display: flex;
        align-items: center;
        justify-content: center;
        overflow: hidden;
      }
      #${PANEL_ID} .announcement-image img {
        width: 100%;
        height: 100%;
        object-fit: cover;
      }
      #${PANEL_ID} .announcement-content {
        flex: 1;
        display: flex;
        flex-direction: column;
        gap: 12px;
      }
      #${PANEL_ID} .announcement-date {
        margin: 0;
        font-size: 12px;
        letter-spacing: 0.08em;
        color: #64748b;
        text-transform: uppercase;
      }
      #${PANEL_ID} #announcement-title {
        margin: 0;
        font-size: 28px;
        font-weight: 700;
        color: #0f172a;
      }
      #${PANEL_ID} .announcement-subtitle {
        margin: 0;
        font-size: 18px;
        font-weight: 600;
        color: #1e293b;
      }
      #${PANEL_ID} .announcement-text {
        margin: 0;
        font-size: 14px;
        line-height: 1.6;
        color: #475569;
      }
      #${PANEL_ID} .announcement-close {
        position: absolute;
        top: 16px;
        right: 16px;
        width: 32px;
        height: 32px;
        border-radius: 50%;
        background: rgba(148, 163, 184, 0.15);
        color: #475569;
        font-size: 18px;
        line-height: 1;
        display: flex;
        align-items: center;
        justify-content: center;
        border: none;
        box-shadow: none;
      }
      #${PANEL_ID} .announcement-close:hover {
        background: rgba(59, 130, 246, 0.18);
        color: ${PRIMARY_COLOR};
      }
      #${PANEL_ID} .announcement-action {
        align-self: center;
        min-width: 160px;
        border-radius: 999px;
        font-weight: 600;
        box-shadow: none;
      }
      @media (max-width: 680px) {
        #${PANEL_ID} .announcement-modal-body {
          flex-direction: column;
        }
        #${PANEL_ID} .announcement-image {
          width: 100%;
          flex: 0 0 auto;
          height: 180px;
        }
      }
      #${PANEL_ID} #pptx-upload-modal {
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,0.5);
        display: none;
        align-items: center;
        justify-content: center;
        z-index: 10002;
        padding: 24px;
        box-sizing: border-box;
      }
      #${PANEL_ID} #pptx-upload-modal[aria-hidden="false"] {
        display: flex;
      }
      #${PANEL_ID} #pptx-upload-modal .csv-modal-dialog {
        background: white;
      }
      #${PANEL_ID} #templates-modal {
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,0.5);
        display: none;
        align-items: center;
        justify-content: center;
        z-index: 10002;
        padding: 24px;
        box-sizing: border-box;
      }
      #${PANEL_ID} #templates-modal[aria-hidden="false"] {
        display: flex;
      }
      #${PANEL_ID} #templates-modal .csv-modal-dialog {
        background: white;
        border-radius: 8px;
        padding: 24px;
        max-height: 90vh;
        overflow-y: auto;
        position: relative;
      }
      #${PANEL_ID} #templates-modal .csv-modal-close {
        position: absolute;
        top: 16px;
        right: 16px;
        background: none;
        border: none;
        font-size: 28px;
        line-height: 1;
        color: #999;
        cursor: pointer;
        padding: 0;
        width: 32px;
        height: 32px;
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 10;
      }
      #${PANEL_ID} #templates-modal .csv-modal-close:hover {
        color: #333;
        background: rgba(0,0,0,0.05);
        border-radius: 4px;
      }
      #${PANEL_ID} #pptx-drop-area {
        border: 2px dashed #cbd5e0;
        border-radius: 8px;
        padding: 40px 20px;
        text-align: center;
        background: #f7fafc;
        margin-bottom: 20px;
        transition: all 0.2s;
      }
      #${PANEL_ID} #pptx-drop-area.drag-over {
        border-color: ${PRIMARY_COLOR};
        background: #e6f2ff;
      }
      #${PANEL_ID} #pptx-drop-area .csv-drop-description {
        font-size: 14px;
        color: #4a5568;
        margin-bottom: 16px;
      }
      #${PANEL_ID} #pptx-file-select-btn {
        padding: 10px 24px;
        background: ${PRIMARY_COLOR};
        color: white;
        border: none;
        border-radius: 6px;
        cursor: pointer;
        font-size: 14px;
        font-weight: 500;
        transition: background 0.2s;
      }
      #${PANEL_ID} #pptx-file-select-btn:hover {
        background: #a50000;
      }
      #${PANEL_ID} #pptx-selected-file {
        padding: 12px;
        background: #f7fafc;
        border-radius: 6px;
        margin-bottom: 20px;
        font-size: 13px;
        color: #2d3748;
      }
      #${PANEL_ID} #progress-indicator .pptx-progress-message {
        font-size: 12px;
        color: #333;
        margin-bottom: 4px;
      }
      #${PANEL_ID} #progress-indicator .pptx-progress-bar {
        width: 80%;
        height: 8px;
        background: #eee;
        margin-bottom: 8px;
      }
      #${PANEL_ID} #progress-indicator .pptx-progress {
        height: 100%;
        width: 0;
        background: ${PRIMARY_COLOR};
        transition: width 0.2s;
      }
      #${PANEL_ID} #progress-indicator .pptx-progress-text {
        font-size: 12px;
        color: #333;
      }
      #${PANEL_ID} #progress-indicator .pptx-cancel {
        margin-top: 8px;
      }
      #${CHAT_CONTAINER_ID} {
        position: fixed;
        top: 0;
        left: 0;
        width: calc(100% - ${PANEL_W});
        height: 100vh;
        background: #fff;
        box-shadow: 2px 0 10px rgba(0,0,0,.2);
        z-index: 9998;
        display: none;
      }
      #${CHAT_CONTAINER_ID} iframe {
        width: 100%;
        height: 100%;
        border: none;
      }
      #${PANEL_ID} #start-text {
        margin-top: 10px;
        background: #fff;
        color: #000;
        color-scheme: light;
      }
      #${PANEL_ID} #start-text::placeholder {
        color: #bbb;
      }
      #${PANEL_ID} #prompt-length-warning {
        color: #d8000c;
        background: #fff4e5;
        border: 1px solid #d8000c;
        border-radius: 4px;
        font-size: 12px;
        padding: 4px 8px;
        margin-top: 4px;
        display: none;
      }
      #${PANEL_ID} .preview-footer {
        padding: 8px;
        border-top: 1px solid #ccc;
        background: #eee;
        display: flex;
        flex-wrap: wrap;
        justify-content: flex-end;
        gap: 8px;
        align-items: center;
        position: relative;
      }
      #${PANEL_ID} #share-message {
        position: absolute;
        bottom: calc(100% + 4px);
        left: 0;
        width: 100%;
        text-align: center;
        font-size: 12px;
        color: #333;
        display: none;
        z-index: 3;
      }
      #${PANEL_ID} #prompt-candidates {
        position: absolute;
        left: 8px;
        right: 8px;
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 4px;
        z-index: 10;
        pointer-events: auto;
        background: rgba(255,255,255,0.9);
        padding: 6px 8px;
        border-radius: 8px;
      }
      #${PANEL_ID} #prompt-candidates.collapsed .prompt-options {
        display: none;
      }
      #${PANEL_ID} #prompt-candidates .prompt-header {
        width: 100%;
        display: flex;
        align-items: center;
        pointer-events: auto;
      }
      #${PANEL_ID} #prompt-candidates .prompt-label {
        font-size: 11px;
        color: #666;
      }
      #${PANEL_ID} #prompt-candidates .prompt-toggle {
        margin-left: auto;
        font-size: 12px;
        color: #666;
        cursor: pointer;
        padding-left: 8px;
      }
      #${PANEL_ID} #prompt-candidates .prompt-options {
        display: flex;
        justify-content: center;
        gap: 8px;
        pointer-events: none;
      }
      #${PANEL_ID} #prompt-candidates .prompt-option {
        pointer-events: auto;
        border: 1px solid transparent;
        border-radius: 8px;
        padding: 4px 8px;
        font-size: 12px;
        cursor: pointer;
        box-shadow: 0 1px 2px rgba(0,0,0,0.1);
        white-space: pre-wrap;
        background: linear-gradient(#fff, #fff) padding-box,
                    var(--border-gradient) border-box;
        transition: background 0.2s;
      }
      #${PANEL_ID} #prompt-candidates .prompt-option.option-1 {
        --border-gradient: linear-gradient(45deg, #ffb3ba, #ffdfba);
      }
      #${PANEL_ID} #prompt-candidates .prompt-option.option-1:hover {
        background: linear-gradient(45deg, rgba(255,179,186,0.2), rgba(255,223,186,0.2)) padding-box,
                    var(--border-gradient) border-box;
      }
      #${PANEL_ID} #prompt-candidates .prompt-option.option-2 {
        --border-gradient: linear-gradient(45deg, #bdecb6, #d6f5c3);
      }
      #${PANEL_ID} #prompt-candidates .prompt-option.option-2:hover {
        background: linear-gradient(45deg, rgba(189,236,182,0.2), rgba(214,245,195,0.2)) padding-box,
                    var(--border-gradient) border-box;
      }
      #${PANEL_ID} #prompt-candidates .prompt-option.option-3 {
        --border-gradient: linear-gradient(45deg, #bae1ff, #d8baff);
      }
      #${PANEL_ID} #prompt-candidates .prompt-option.option-3:hover {
        background: linear-gradient(45deg, rgba(186,225,255,0.2), rgba(216,186,255,0.2)) padding-box,
                    var(--border-gradient) border-box;
      }
      #${PANEL_ID} #prompt-candidates .prompt-option.option-4 {
        --border-gradient: linear-gradient(45deg, #fff3b0, #ffe066);
      }
      #${PANEL_ID} #prompt-candidates .prompt-option.option-4:hover {
        background: linear-gradient(45deg, rgba(255,243,176,0.3), rgba(255,224,102,0.3)) padding-box,
                    var(--border-gradient) border-box;
      }
      #${PANEL_ID} .preview-container {
        width: 100%;
        box-sizing: border-box;
      }
      #${PANEL_ID} .preview-container.multi-slide {
        overflow-y: auto;
        overflow-x: hidden;
        max-height: calc(100vh - 160px);
        padding-right: 4px;
        scrollbar-width: thin;
      }
      #${PANEL_ID} .preview-container.multi-slide::-webkit-scrollbar {
        width: 6px;
      }
      #${PANEL_ID} .preview-container.multi-slide::-webkit-scrollbar-track {
        background: transparent;
      }
      #${PANEL_ID} .preview-container.multi-slide::-webkit-scrollbar-thumb {
        background: rgba(0,0,0,0.2);
        border-radius: 3px;
      }
      #${PANEL_ID} .preview-slide {
        margin-bottom: 16px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.2);
        overflow: hidden;
      }
      #${PANEL_ID} .json-slide-title {
        font-size: 56px;
        font-weight: bold;
        margin-bottom: 24px;
        line-height: 1.2;
        text-align: left;
      }
      #${PANEL_ID} .json-slide-prompt {
        font-size: 24px;
        white-space: pre-wrap;
        line-height: 1.4;
        text-align: left;
      }
      #start-arrow {
        width: 24px;
        height: 24px;
        position: fixed;
        pointer-events: none;
        z-index: 10000;
        animation: arrow-bounce 0.8s infinite;
      }
      @keyframes arrow-bounce {
        0%, 100% { transform: translateY(0); }
        50% { transform: translateY(-6px); }
      }
      @keyframes spin {
        from { transform: rotate(0deg); }
        to { transform: rotate(360deg); }
      }
      @keyframes glitter {
        0%, 100% { box-shadow: 0 0 4px ${PRIMARY_COLOR}; }
        50% { box-shadow: 0 0 12px 4px ${PRIMARY_COLOR}; }
      }
      /* PPTX Preview styles in preview-container */
      #${PANEL_ID} .preview-container .pptx-preview-wrapper {
        background: #f8f9fa;
        border: 2px solid #e9ecef;
        border-radius: 8px;
        padding: 0;
        width: 100%;
        aspect-ratio: 16 / 9;
        max-height: calc(100vh - 200px);
        display: flex;
        align-items: center;
        justify-content: center;
        position: relative;
        overflow: hidden;
        margin: 0 auto;
      }
      #${PANEL_ID} .preview-container .pptx-preview-wrapper > * {
        width: 100% !important;
        height: 100% !important;
        max-width: 100%;
        max-height: 100%;
      }
      #${PANEL_ID} .preview-container .pptx-preview-wrapper .slide {
        width: 100% !important;
        height: 100% !important;
        object-fit: contain;
      }
      #${PANEL_ID} .preview-container .pptx-preview-wrapper svg {
        width: 100% !important;
        height: 100% !important;
        max-width: 100%;
        max-height: 100%;
      }
      #${PANEL_ID} .preview-container .pptx-preview-wrapper svg text:empty,
      #${PANEL_ID} .preview-container .pptx-preview-wrapper svg text[text-anchor]:empty {
        display: none;
      }
      #${PANEL_ID} .preview-container .pptx-preview-wrapper .pptx-preview-wrapper-inner {
        width: 100% !important;
        height: 100% !important;
        display: flex;
        align-items: center;
        justify-content: center;
      }
      #${PANEL_ID} .preview-container .pptx-preview-wrapper .pptx-preview-wrapper-inner > div {
        max-width: 100%;
        max-height: 100%;
      }
      .preview-error {
        text-align: center;
        color: #d32f2f;
        padding: 20px;
      }
      .preview-error p {
        margin: 8px 0;
      }
    `;
    const styleEl = Object.assign(document.createElement('style'), {
      id: STYLE_ID,
      textContent: css,
    });
    document.head.appendChild(styleEl);
  }

  function injectToggle() {
    const { TOGGLE_ID, PREVIEW_PARAM, HTML_PARAM, HTML_LIST_PARAM } = app;
    if (document.getElementById(TOGGLE_ID)) return;
    injectStyles();
    const button = Object.assign(document.createElement('button'), {
      id: TOGGLE_ID,
      type: 'button',
    });
    const img = document.createElement('img');
    img.src = safeGetURL('images/Arrow.png');
    button.appendChild(img);
    button.onclick = togglePanel;
    document.body.appendChild(button);
    try {
      const p = app.payload || {};
      const htmlParam = p[HTML_PARAM];
      if (htmlParam) {
        const html = LZString.decompressFromEncodedURIComponent(htmlParam);
        if (html) app.lastHtml = html;
      }
      const listParam = p[HTML_LIST_PARAM];
      if (Array.isArray(listParam)) {
        app.htmlSlides = listParam
          .map(h => {
            try { return LZString.decompressFromEncodedURIComponent(h); } catch { return ''; }
          })
          .filter(Boolean);
        if (app.htmlSlides.length > 1) app.multiSlideMode = true;
      }
      if (p[PREVIEW_PARAM] === '1') {
        // auto open removed; user must click the toggle to open
      }
    } catch (e) {}
    if (app.panelOpen) {
      button.classList.add('open');
      openPanel();
    }
  }

  function togglePanel() {
    const { TOGGLE_ID } = app;
    const isOpen = !!document.getElementById(app.PANEL_ID);
    const toggle = document.getElementById(TOGGLE_ID);
    if (toggle) toggle.classList.toggle('open', !isOpen);
    isOpen ? closePanel() : openPanel();
  }

  function closePanel() {
    const { PANEL_ID, ACTIVE_CLS } = app;
    document.getElementById(PANEL_ID)?.remove();
    document.body.classList.remove(ACTIVE_CLS);
    app.panelOpen = false;
    savePanelState(app);
    resetSearchMode(true);
    if (app.observer) {
      app.observer.disconnect();
      app.observer = null;
    }
    clearTimeout(app.renderTimer);
    app.lastPptx = '';
    app.scrapedCode = '';
    app.lastHtml = '';
    app.htmlSlides = [];
    app.startText = '';
    app.templateModalOpen = false;
    app.templateHtml = '';
    app.isTyping = false;
    app.pendingRender = false;
    app.showHome = false;
    app.csvPrompts = [];
    app.csvRows = [];
    app.csvFileName = '';
    if (app.chatFrame) {
      app.chatFrame.remove();
      app.chatFrame = null;
    }
    app.detectedHtml = false;
    const chat = document.getElementById(app.CHAT_CONTAINER_ID);
    if (chat) chat.style.display = 'none';
  }

  function openPanel() {
    const { PANEL_ID, ACTIVE_CLS, BOX_ID } = app;
    document.body.classList.add(ACTIVE_CLS);
    app.panelOpen = true;
    savePanelState(app);

    const panel = document.createElement('div');
    const announcementCopy = getAnnouncementCopy(currentLang);
    const announcementBodyHtml = (announcementCopy?.body || [])
      .map((line) => `<p class="announcement-text">${line}</p>`)
      .join('');
    const announcementImage = safeGetURL(
      CURRENT_ANNOUNCEMENT?.imageSrc || 'images/Explanation.svg'
    );
    panel.id = PANEL_ID;
    panel.innerHTML = `
      <div class="preview-header">
        <div>
          <button id="home-btn" type="button"><img src="${safeGetURL(app.HOME_GRAY_SRC)}" style="width:14px;height:14px;"></button>
          <button id="convert-btn" type="button" disabled><img src="${safeGetURL(app.CONVERT_ICON_SRC)}" style="width:14px;height:14px;margin-right:4px;"><span data-i18n="convert"></span></button>
          <!-- <button id="copy-html-btn" type="button" style="display:none;" disabled>Copy HTML</button> -->
          <button id="preview-pptx-btn" type="button" disabled style="display:none;"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 576 512" style="width:14px;height:14px;margin-right:4px;fill:white;"><path d="M572.52 241.4C518.29 135.59 410.93 64 288 64S57.68 135.64 3.48 241.41a32.35 32.35 0 0 0 0 29.19C57.71 376.41 165.07 448 288 448s230.32-71.64 284.52-177.41a32.35 32.35 0 0 0 0-29.19zM288 400a144 144 0 1 1 144-144 143.93 143.93 0 0 1-144 144zm0-240a95.31 95.31 0 0 0-25.31 3.79 47.85 47.85 0 0 1-66.9 66.9A95.78 95.78 0 1 0 288 160z"/></svg><span data-i18n="previewPptx"></span></button>
          <button id="save-preview-btn" type="button" disabled><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 448 512" style="width:14px;height:14px;margin-right:4px;fill:white;"><path d="M433.941 129.941l-83.882-83.882A48 48 0 0 0 316.118 32H48C21.49 32 0 53.49 0 80v352c0 26.51 21.49 48 48 48h352c26.51 0 48-21.49 48-48V163.882a48 48 0 0 0-14.059-33.941zM224 416c-35.346 0-64-28.654-64-64 0-35.346 28.654-64 64-64s64 28.654 64 64c0 35.346-28.654 64-64 64zm96-304.52V212c0 6.627-5.373 12-12 12H76c-6.627 0-12-5.373-12-12V108c0-6.627 5.373-12 12-12h228.52c3.183 0 6.235 1.264 8.485 3.515l3.48 3.48A11.996 11.996 0 0 1 320 111.48z"/></svg><span data-i18n="savePreview"></span></button>
          <button id="download-btn" type="button" disabled><img src="${safeGetURL(app.DOWNLOAD_ICON_SRC)}" style="width:14px;height:14px;margin-right:4px;"><span data-i18n="export"></span></button>
          <button id="download-api-btn" type="button" disabled><img src="${safeGetURL(app.DOWNLOAD_ICON_SRC)}" style="width:14px;height:14px;margin-right:4px;"><span data-i18n="downloadViaApi"></span></button>
          <button id="download-multi-btn" type="button" data-i18n="exportMulti"></button>
          <button id="start-all-btn" type="button" style="display:none;" data-i18n="startSlides"></button>
          <button id="create-pptx-btn" type="button" data-i18n="createPptx"></button>
        </div>
        <div class="toggle-wrap">
          <label class="switch">
            <input type="checkbox" id="multiToggle">
            <span class="slider"></span>
          </label>
          <span class="toggle-label" data-i18n="multiSlide"></span>
          <button id="api-settings-btn" type="button" title="API Settings" style="margin-left:8px;"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" style="width:16px;height:16px;fill:currentColor;"><path d="M487.4 315.7l-42.6-24.6c4.3-23.2 4.3-47 0-70.2l42.6-24.6c4.9-2.8 7.1-8.6 5.5-14-11.1-35.6-30-67.8-54.7-94.6-3.8-4.1-10-5.1-14.8-2.3L380.8 110c-17.9-15.4-38.5-27.3-60.8-35.1V25.8c0-5.6-3.9-10.5-9.4-11.7-36.7-8.2-74.3-7.8-109.2 0-5.5 1.2-9.4 6.1-9.4 11.7V75c-22.2 7.9-42.8 19.8-60.8 35.1L88.7 85.5c-4.9-2.8-11-1.9-14.8 2.3-24.7 26.7-43.6 58.9-54.7 94.6-1.7 5.4.6 11.2 5.5 14L67.3 221c-4.3 23.2-4.3 47 0 70.2l-42.6 24.6c-4.9 2.8-7.1 8.6-5.5 14 11.1 35.6 30 67.8 54.7 94.6 3.8 4.1 10 5.1 14.8 2.3l42.6-24.6c17.9 15.4 38.5 27.3 60.8 35.1v49.2c0 5.6 3.9 10.5 9.4 11.7 36.7 8.2 74.3 7.8 109.2 0 5.5-1.2 9.4-6.1 9.4-11.7v-49.2c22.2-7.9 42.8-19.8 60.8-35.1l42.6 24.6c4.9 2.8 11 1.9 14.8-2.3 24.7-26.7 43.6-58.9 54.7-94.6 1.5-5.5-.7-11.3-5.6-14.1zM256 336c-44.1 0-80-35.9-80-80s35.9-80 80-80 80 35.9 80 80-35.9 80-80 80z"/></svg></button>
        </div>
      </div>
        <div id="preview-content">
        <div class="preview-container"></div>
        <div id="progress-indicator">
          <div class="pptx-progress-message" data-i18n="progressPreparingExport"></div>
          <div class="pptx-progress-bar"><div class="pptx-progress"></div></div>
          <div class="pptx-progress-text">0%</div>
          <button type="button" class="pptx-cancel" data-i18n="cancel"></button>
        </div>
        <div id="wait-overlay">
          <div class="wait-spinner"></div>
          <div class="wait-message" data-i18n="pleaseWait"></div>
        </div>
        <div id="pptx-detected-overlay">
          <div data-i18n="exportReady"></div>
          <div class="pptx-overlay-actions">
            <button type="button" id="pptx-only-download" data-i18n="exportPptx"></button>
            <button type="button" id="pptx-fast-download" data-i18n="downloadViaApi"></button>
          </div>
        </div>
        </div>
        <div id="announcement-modal" aria-hidden="true">
          <div class="announcement-modal-dialog" role="dialog" aria-modal="true" aria-labelledby="announcement-title">
            <button type="button" class="announcement-close" id="announcement-close" aria-label="Close">×</button>
            <div class="announcement-modal-body">
              <div class="announcement-image">
                <img src="${announcementImage}" alt="" role="presentation">
              </div>
              <div class="announcement-content">
                <p class="announcement-date" data-announcement-slot="date">${announcementCopy?.date || ''}</p>
                <h2 id="announcement-title" data-announcement-slot="version">${announcementCopy?.version || ''}</h2>
                <h3 class="announcement-subtitle" data-announcement-slot="subtitle">${announcementCopy?.subtitle || ''}</h3>
                <div class="announcement-body" data-announcement-body>${announcementBodyHtml}</div>
              </div>
            </div>
            <button type="button" id="announcement-confirm" class="announcement-action">${announcementCopy?.actionLabel || ''}</button>
          </div>
        </div>
        <div id="pptx-upload-modal" aria-hidden="true" style="display:none;">
          <div class="csv-modal-dialog">
            <button type="button" class="csv-modal-close pptx-modal-close" aria-label="Close">×</button>
            <div id="pptx-drop-area">
              <p class="csv-drop-description" data-i18n="pptxDropDescription"></p>
              <button type="button" id="pptx-file-select-btn" data-i18n="pptxSelectFile"></button>
              <input type="file" id="pptx-file-input" accept=".pptx" style="display:none">
            </div>
            <div id="pptx-selected-file" data-i18n="pptxNoFile"></div>
            <div id="pptx-analysis-result" style="display:none;">
              <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
                <h3 style="margin:0;" data-i18n="pptxAnalysisTitle"></h3>
                <div style="display:flex;gap:8px;">
                  <button type="button" id="pptx-copy-btn" style="display:none;padding:6px 16px;background:#f5f5f5;color:#333;border:1px solid #ddd;border-radius:4px;cursor:pointer;font-size:13px;" data-i18n="pptxCopyJson"></button>
                  <button type="button" id="pptx-send-btn" style="display:none;padding:6px 16px;background:#bf0000;color:white;border:none;border-radius:4px;cursor:pointer;font-size:13px;font-weight:500;" data-i18n="pptxSendJson"></button>
                </div>
              </div>
              <div id="pptx-slide-list"></div>
              <textarea id="pptx-json-output" readonly style="width:100%;height:400px;font-family:monospace;font-size:11px;white-space:pre;overflow:auto;border:1px solid #ddd;border-radius:4px;padding:8px;box-sizing:border-box;"></textarea>
              <div style="margin-top:6px;font-size:11px;color:#666;"><em>ℹ️ 表示エリアをスクロールして全データを確認できます。コピーボタンで全データがクリップボードにコピーされます。</em></div>
            </div>
            <div class="csv-modal-actions">
              <div class="csv-modal-actions-left"></div>
              <div class="csv-modal-actions-right">
                <button type="button" id="pptx-cancel-btn" data-i18n="cancel"></button>
              </div>
            </div>
          </div>
        </div>
        <div id="api-settings-modal" aria-hidden="true" style="display:none;">
          <div class="csv-modal-dialog">
            <button type="button" class="csv-modal-close api-modal-close" aria-label="Close">×</button>
            <h2 style="margin-bottom:20px;" data-i18n="apiKeySettings"></h2>
            <div style="margin-bottom:15px;">
              <label style="display:block;margin-bottom:5px;font-weight:bold;" data-i18n="apiKeyCurrentStatus"></label>
              <div id="api-key-status" style="padding:8px;background:#f5f5f5;border-radius:4px;"></div>
            </div>
            <div style="margin-bottom:15px;">
              <label style="display:block;margin-bottom:5px;font-weight:bold;">現在のIPアドレス:</label>
              <div id="api-current-ip-status" style="padding:8px;background:#f5f5f5;border-radius:4px;">クリックして確認</div>
              <button type="button" id="api-check-ip-btn" style="margin-top:8px;padding:6px 12px;background:#007bff;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;">IPアドレスを確認</button>
            </div>
            <div style="margin-bottom:20px;">
              <label for="api-key-input" style="display:block;margin-bottom:5px;font-weight:bold;">API Key:</label>
              <input type="password" id="api-key-input" placeholder="" data-i18n-placeholder="apiKeyPlaceholder" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:4px;font-family:monospace;">
            </div>
            <div class="csv-modal-actions">
              <div class="csv-modal-actions-left">
                <button type="button" id="api-key-clear-btn" data-i18n="apiKeyClear"></button>
              </div>
              <div class="csv-modal-actions-right">
                <button type="button" id="api-key-cancel-btn" data-i18n="apiKeyCancel"></button>
                <button type="button" id="api-key-save-btn" data-i18n="apiKeySave"></button>
              </div>
            </div>
          </div>
        </div>
        <div id="templates-modal" aria-hidden="true" style="display:none;">
          <div class="csv-modal-dialog" style="max-width:900px;">
            <button type="button" class="csv-modal-close templates-modal-close" aria-label="Close">×</button>
            <h2 style="margin-top:0;margin-bottom:20px;" data-i18n="templatesTitle"></h2>
            <div id="templates-list"></div>
          </div>
        </div>
        <div class="preview-footer">
          <div id="share-message"></div>
          <button id="inquiry-btn" type="button" data-i18n="inquiry"></button>
          <button id="manual-btn" type="button" data-i18n="manual"></button>
          <button id="share-btn" type="button" data-i18n="share"></button>
          <div id="lang-selector" class="dropup">
            <button id="lang-btn"></button>
            <ul id="lang-menu">
              <li data-lang="ja" data-i18n="langJa"></li>
              <li data-lang="en" data-i18n="langEn"></li>
            </ul>
          </div>
        </div>
      `;
    document.body.appendChild(panel);

    panel.querySelector('#download-btn').onclick    = () => app.downloadPptx && app.downloadPptx();

    // API経由ダウンロードボタンのイベントリスナー
    const downloadApiBtn = panel.querySelector('#download-api-btn');
    if (downloadApiBtn) {
      downloadApiBtn.onclick = async () => {
        try {
          if (!app.scrapedCode) {
            alert(t('noTargetCode'));
            return;
          }

          // APIクライアントを動的にインポート
          const { generatePptxViaApi, getApiKey } = await import(chrome.runtime.getURL('src/apiClient.js'));

          // APIキーの確認
          const apiKey = await getApiKey();
          if (!apiKey) {
            alert(t('apiKeyNotSet'));
            return;
          }

          // 進捗表示を開始
          app.showProgress(t('apiGenerating'), 120000);

          // ファイル名を生成
          let fileName = 'presentation.pptx';
          try {
            const m = app.scrapedCode.match(/addText\(\s*(["'`])([\s\S]*?)\1/);
            if (m) {
              const text = m[2].replace(/\\n/g, ' ').trim();
              // 日本語を含むファイル名をサニタイズ
              // 不正な文字を削除し、20文字に制限
              let sanitized = text.replace(/[\\/:*?"<>|]/g, '_').slice(0, 20);

              // 日本語が含まれる場合、そのまま使用（ブラウザが適切にエンコード）
              // ただし、空白やタブは_に置換
              sanitized = sanitized.replace(/\s+/g, '_');

              const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
              fileName = `${sanitized}_${date}.pptx`;
            }
          } catch (e) {
            console.error('[API Download] ファイル名生成エラー', e);
          }

          // API経由でパワーポイントを生成
          await generatePptxViaApi(app.scrapedCode, fileName);

          // 進捗表示を終了
          app.stopFakeProgress();
          app.updateProgress(100);
          setTimeout(() => {
            app.hideProgress();
            // alert(t('apiDownloadSuccess')); // アラート削除：ダウンロード成功は通知不要
          }, 500);

        } catch (error) {
          console.error('[API Download] エラー:', error);
          app.stopFakeProgress();
          app.hideProgress();

          // エラーメッセージの判定
          let errorMsg = t('apiDownloadFailed');
          if (error.message.includes('APIキーが設定されていません')) {
            errorMsg = t('apiKeyNotSet');
          } else if (error.message.includes('APIキーが無効')) {
            errorMsg = t('apiKeyInvalid');
          } else {
            errorMsg = `${t('apiDownloadFailed')}: ${error.message}`;
          }

          alert(errorMsg);
        }
      };
    }

    const multiBtn = panel.querySelector('#download-multi-btn');
    if (multiBtn) multiBtn.onclick = downloadMultiplePptx;
    const createBtn = panel.querySelector('#create-pptx-btn');
    if (createBtn) createBtn.onclick = startHtmlSlides;
    const startAllBtn = panel.querySelector('#start-all-btn');
    if (startAllBtn) startAllBtn.onclick = startBatchCreation;
    const copyBtnEl = panel.querySelector('#copy-html-btn');
    if (copyBtnEl) copyBtnEl.onclick = copyHtml;
    panel.querySelector('#convert-btn').onclick     = toggleChat;
    panel.querySelector('#home-btn').onclick        = toggleHome;
    panel.querySelector('#progress-indicator .pptx-cancel').onclick = cancelDownload;

    // PPTX Preview event listeners
    const previewPptxBtn = panel.querySelector('#preview-pptx-btn');
    if (previewPptxBtn) {
      console.log('[Preview] Setting up preview button event listener');
      previewPptxBtn.onclick = showPptxPreview;
    } else {
      console.warn('[Preview] preview-pptx-btn not found in panel');
    }

    // Save Preview button event listener
    const savePreviewBtn = panel.querySelector('#save-preview-btn');
    if (savePreviewBtn) {
      savePreviewBtn.onclick = saveCodeAsTemplate;
    }
    const inquiryBtn = panel.querySelector('#inquiry-btn');
    if (inquiryBtn) inquiryBtn.onclick = () => {
      window.open(app.INQUIRY_URL, '_blank', 'noopener');
    };
    const manualBtn = panel.querySelector('#manual-btn');
    if (manualBtn) manualBtn.onclick = () => {
      window.open(app.MANUAL_URL, '_blank', 'noopener');
    };
    const shareBtn = panel.querySelector('#share-btn');
    const shareMsg = panel.querySelector('#share-message');
    if (shareBtn) shareBtn.onclick = async () => {
      const copied = await copyToClipboard(app.SHARE_URL);
      if (copied && shareMsg) {
        shareMsg.dataset.i18n = 'linkCopied';
        shareMsg.textContent = t('linkCopied');
        shareMsg.style.display = 'block';
      }
      const orig = shareBtn.dataset.i18n || 'share';
      shareBtn.dataset.i18n = copied ? 'copied' : 'copyFailed';
      shareBtn.textContent = t(shareBtn.dataset.i18n);
      setTimeout(() => {
        if (shareMsg) shareMsg.style.display = 'none';
        shareBtn.dataset.i18n = orig;
        shareBtn.textContent = t(orig);
      }, 3000);
    };

    const langBtn = panel.querySelector('#lang-btn');
    const langSelector = panel.querySelector('#lang-selector');
    const langMenu = panel.querySelector('#lang-menu');
    if (langBtn && langMenu && langSelector) {
      langBtn.onclick = () => langSelector.classList.toggle('open');
      langMenu.querySelectorAll('li').forEach((li) => {
        li.onclick = () => {
          setLanguage(li.dataset.lang);
          langSelector.classList.remove('open');
        };
      });
    }


    // PPTX Upload Modal
    const pptxModal = panel.querySelector('#pptx-upload-modal');
    const pptxClose = panel.querySelector('.pptx-modal-close');
    const pptxCancel = panel.querySelector('#pptx-cancel-btn');
    const pptxInput = panel.querySelector('#pptx-file-input');
    const pptxSelectBtn = panel.querySelector('#pptx-file-select-btn');
    const pptxDropArea = panel.querySelector('#pptx-drop-area');
    const pptxSelectedFile = panel.querySelector('#pptx-selected-file');
    const pptxAnalysisResult = panel.querySelector('#pptx-analysis-result');
    const pptxSlideList = panel.querySelector('#pptx-slide-list');
    const pptxJsonOutput = panel.querySelector('#pptx-json-output');
    const pptxCopyBtn = panel.querySelector('#pptx-copy-btn');
    const pptxSendBtn = panel.querySelector('#pptx-send-btn');

    if (pptxModal && pptxClose && pptxCancel && pptxInput &&
        pptxSelectBtn && pptxDropArea && pptxSelectedFile && pptxAnalysisResult &&
        pptxSlideList && pptxJsonOutput && pptxCopyBtn && pptxSendBtn) {

      // Store openPptxModal in app so it can be used from renderPreview
      app.openPptxModal = () => {
        pptxModal.setAttribute('aria-hidden', 'false');
        pptxModal.style.display = 'flex';
      };

      const closePptxModal = () => {
        pptxModal.setAttribute('aria-hidden', 'true');
        pptxModal.style.display = 'none';
        pptxInput.value = '';
        pptxSelectedFile.dataset.i18n = 'pptxNoFile';
        applyTranslations(pptxSelectedFile);
        pptxAnalysisResult.style.display = 'none';
        pptxCopyBtn.style.display = 'none';
        pptxSendBtn.style.display = 'none';
        pptxJsonOutput.value = '';
        pptxSlideList.innerHTML = '';
      };

      const processPptxFile = async (file) => {
        if (!file) return;

        pptxSelectedFile.textContent = file.name;
        pptxSelectedFile.removeAttribute('data-i18n');
        pptxAnalysisResult.style.display = 'block';
        pptxSlideList.innerHTML = '<p>解析中...</p>';
        pptxJsonOutput.value = '';
        pptxCopyBtn.style.display = 'none';
        pptxSendBtn.style.display = 'none';

        try {
          // Load pptxAnalyzer module
          const { analyzePPTX } = await import(chrome.runtime.getURL('src/pptxAnalyzer.js'));

          const result = await analyzePPTX(file);

          if (result.success) {
            pptxSlideList.innerHTML = `<p>✅ ${result.slideCount}枚のスライドを解析しました</p>`;

            // Show slide selector
            const select = document.createElement('select');
            select.style.cssText = 'width:100%;padding:8px;margin:10px 0;font-size:14px;';
            result.slides.forEach(slide => {
              const option = document.createElement('option');
              option.value = slide.slideNumber - 1;
              option.textContent = `スライド ${slide.slideNumber} (要素:${slide.elementCount}, 表:${slide.tableCount}, 線:${slide.lineCount})`;
              select.appendChild(option);
            });

            // Create info display for data statistics
            const infoDiv = document.createElement('div');
            infoDiv.style.cssText = 'margin:10px 0;padding:8px;background:#f0f7ff;border:1px solid #b3d9ff;border-radius:4px;font-size:12px;color:#333;';
            infoDiv.id = 'pptx-data-info';

            const updateJsonOutput = () => {
              const selectedIndex = parseInt(select.value);
              const promptText = result.slides[selectedIndex].promptWithJson;
              pptxJsonOutput.value = promptText;

              // Calculate and display data statistics
              const charCount = promptText.length;
              const tableMatches = promptText.match(/"tables":\s*\[/g);
              const hasTable = tableMatches && tableMatches.length > 0;

              let totalRows = 0;
              if (hasTable) {
                // Count total rows across all tables by counting "h": properties in rows array
                const rowsMatches = promptText.match(/"rows":\s*\[([\s\S]*?)\]/g);
                if (rowsMatches) {
                  rowsMatches.forEach(rowsBlock => {
                    const rowHeights = rowsBlock.match(/"h":\s*[\d.]+/g);
                    if (rowHeights) totalRows += rowHeights.length;
                  });
                }
              }

              infoDiv.innerHTML = `📊 データ統計: <strong>${charCount.toLocaleString()}</strong> 文字 | テーブル: <strong>${result.slides[selectedIndex].tableCount}</strong> 個 | テーブル総行数: <strong>${totalRows}</strong> 行 ${totalRows > 0 ? '✅ 全行抽出済み' : ''}`;
            };

            select.addEventListener('change', updateJsonOutput);
            pptxSlideList.appendChild(select);
            pptxSlideList.appendChild(infoDiv);

            // Show first slide by default
            updateJsonOutput();
            pptxCopyBtn.style.display = 'inline-block';
            pptxSendBtn.style.display = 'inline-block';
          } else {
            pptxSlideList.innerHTML = `<p style="color:red;">❌ エラー: ${result.error}</p>`;
          }
        } catch (err) {
          pptxSlideList.innerHTML = `<p style="color:red;">❌ エラー: ${err.message}</p>`;
          console.error('PPTX analysis error:', err);
        }
      };

      // pptxTextBtn is set up later in renderPreview() when home screen is shown
      pptxClose.onclick = closePptxModal;
      pptxCancel.onclick = closePptxModal;
      pptxSelectBtn.onclick = () => pptxInput.click();

      pptxInput.onchange = () => {
        const file = pptxInput.files && pptxInput.files[0];
        if (file) processPptxFile(file);
      };

      pptxCopyBtn.onclick = () => {
        pptxJsonOutput.select();
        document.execCommand('copy');
        const originalText = pptxCopyBtn.textContent;
        pptxCopyBtn.textContent = 'コピーしました！';
        setTimeout(() => {
          pptxCopyBtn.textContent = originalText;
        }, 2000);
      };

      pptxSendBtn.onclick = async () => {
        const jsonText = pptxJsonOutput.value;
        if (!jsonText) {
          alert('送信するJSONがありません');
          return;
        }

        // Close the modal
        closePptxModal();

        // Navigate to CHAT_URL (same as "パワポに変換" button) and pass the JSON via handOff
        const convertUrl = app.CHAT_URL;

        // Use handOff to save the prompt and navigate
        await handOff({ prompt: jsonText }, convertUrl);
      };

      // Drag and drop support
      const preventDefaults = (e) => {
        e.preventDefault();
        e.stopPropagation();
      };

      ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        pptxDropArea.addEventListener(eventName, preventDefaults, false);
      });

      ['dragenter', 'dragover'].forEach(eventName => {
        pptxDropArea.addEventListener(eventName, () => {
          pptxDropArea.classList.add('drag-over');
        }, false);
      });

      ['dragleave', 'drop'].forEach(eventName => {
        pptxDropArea.addEventListener(eventName, () => {
          pptxDropArea.classList.remove('drag-over');
        }, false);
      });

      pptxDropArea.addEventListener('drop', (e) => {
        const files = e.dataTransfer && e.dataTransfer.files;
        if (files && files.length) {
          processPptxFile(files[0]);
        }
      }, false);
    }

    // API Settings Modal
    const apiSettingsBtn = panel.querySelector('#api-settings-btn');
    const apiSettingsModal = panel.querySelector('#api-settings-modal');
    const apiModalClose = panel.querySelector('.api-modal-close');
    const apiKeySaveBtn = panel.querySelector('#api-key-save-btn');
    const apiKeyCancelBtn = panel.querySelector('#api-key-cancel-btn');
    const apiKeyClearBtn = panel.querySelector('#api-key-clear-btn');
    const apiKeyInput = panel.querySelector('#api-key-input');
    const apiKeyStatus = panel.querySelector('#api-key-status');

    if (apiSettingsBtn && apiSettingsModal && apiModalClose && apiKeySaveBtn &&
        apiKeyCancelBtn && apiKeyClearBtn && apiKeyInput && apiKeyStatus) {

      const updateApiKeyStatus = async () => {
        const { loadApiKey } = await import(chrome.runtime.getURL('src/storage.js'));
        const apiKey = await loadApiKey();
        if (apiKey) {
          apiKeyStatus.innerHTML = `✅ <span data-i18n="apiKeySet"></span>`;
          applyTranslations(apiKeyStatus);
        } else {
          apiKeyStatus.innerHTML = `⚠️ <span data-i18n="apiKeyNotSetStatus"></span>`;
          applyTranslations(apiKeyStatus);
        }
      };

      const closeApiModal = () => {
        apiSettingsModal.style.display = 'none';
        apiSettingsModal.setAttribute('aria-hidden', 'true');
        apiKeyInput.value = '';
        document.removeEventListener('keydown', apiEscHandler, true);
      };

      const apiEscHandler = (e) => {
        if (e.key === 'Escape') {
          closeApiModal();
        }
      };

      const openApiModal = async () => {
        apiSettingsModal.style.display = 'flex';
        apiSettingsModal.setAttribute('aria-hidden', 'false');
        await updateApiKeyStatus();

        // Load existing API key
        const { loadApiKey } = await import(chrome.runtime.getURL('src/storage.js'));
        const existingKey = await loadApiKey();
        if (existingKey) {
          apiKeyInput.value = existingKey;
        }

        document.addEventListener('keydown', apiEscHandler, true);
        setTimeout(() => apiKeyInput.focus(), 0);
      };

      apiSettingsBtn.onclick = openApiModal;
      apiModalClose.onclick = closeApiModal;
      apiKeyCancelBtn.onclick = closeApiModal;

      apiSettingsModal.addEventListener('click', (e) => {
        if (e.target === apiSettingsModal) {
          closeApiModal();
        }
      });

      apiKeySaveBtn.onclick = async () => {
        const apiKey = apiKeyInput.value.trim();
        if (!apiKey) {
          alert(t('apiKeyInvalid'));
          return;
        }

        try {
          const { saveApiKey } = await import(chrome.runtime.getURL('src/storage.js'));
          await saveApiKey(apiKey);
          alert(t('apiKeySaved'));
          await updateApiKeyStatus();
          closeApiModal();
        } catch (error) {
          console.error('Failed to save API key:', error);
          alert('Failed to save API key');
        }
      };

      apiKeyClearBtn.onclick = async () => {
        if (!confirm(t('apiKeyCleared') + '?')) {
          return;
        }

        try {
          const { clearApiKey } = await import(chrome.runtime.getURL('src/storage.js'));
          await clearApiKey();
          apiKeyInput.value = '';
          alert(t('apiKeyCleared'));
          await updateApiKeyStatus();
        } catch (error) {
          console.error('Failed to clear API key:', error);
          alert('Failed to clear API key');
        }
      };

      // IP確認ボタン
      const apiCheckIpBtn = panel.querySelector('#api-check-ip-btn');
      const apiCurrentIpStatus = panel.querySelector('#api-current-ip-status');
      if (apiCheckIpBtn && apiCurrentIpStatus) {
        apiCheckIpBtn.onclick = async () => {
          try {
            apiCheckIpBtn.disabled = true;
            apiCheckIpBtn.textContent = '確認中...';
            apiCurrentIpStatus.innerHTML = '🔄 確認中...';

            const { checkCurrentIp } = await import(chrome.runtime.getURL('src/apiClient.js'));
            const result = await checkCurrentIp();

            // IPv4とIPv6の両方を表示
            let ipDisplay = '';
            if (result.ipv4 && result.ipv4 !== 'Unknown') {
              ipDisplay += `<strong>IPv4:</strong> ${result.ipv4}<br>`;
            }
            if (result.ipv6 && result.ipv6 !== 'Unknown') {
              ipDisplay += `<strong>IPv6:</strong> ${result.ipv6}<br>`;
            }
            if (!ipDisplay) {
              ipDisplay = 'IPアドレスを取得できませんでした<br>';
            }

            if (result.allowed) {
              apiCurrentIpStatus.innerHTML = `✅ ${ipDisplay}<small style="color:#28a745;font-weight:bold;">このIPアドレスはアクセス許可されています</small>`;
            } else {
              const messageDetail = result.message ? `<br><br><small style="color:#666;">${result.message}</small>` : '';
              apiCurrentIpStatus.innerHTML = `🚫 ${ipDisplay}<small style="color:#dc3545;font-weight:bold;">このIPアドレスはアクセス拒否されています</small>${messageDetail}<br><br><small style="color:#dc3545;">対処法:<br>1. VPNを切断してください（Zscaler/Netskope等）<br>2. Rakuten INTRA社内ネットワークに直接接続してください<br>3. または管理者にこのIPを許可リストに追加してもらってください</small>`;
            }

            apiCheckIpBtn.disabled = false;
            apiCheckIpBtn.textContent = '再確認';
          } catch (error) {
            console.error('Failed to check IP:', error);
            apiCurrentIpStatus.innerHTML = `❌ 確認失敗<br><small style="color:#dc3545;">${error.message}</small>`;
            apiCheckIpBtn.disabled = false;
            apiCheckIpBtn.textContent = 'IPアドレスを確認';
          }
        };
      }
    }

    // Templates Modal
    const templatesBtn = panel.querySelector('#templates-btn');
    const templatesModal = panel.querySelector('#templates-modal');
    const templatesClose = panel.querySelector('.templates-modal-close');

    console.log('[Templates Modal Setup]', { templatesBtn, templatesModal, templatesClose });

    if (templatesBtn && templatesModal && templatesClose) {
      templatesBtn.onclick = () => {
        console.log('[Templates Modal] Templates button clicked');
        openTemplatesModal();
      };

      const closeTemplatesModalHandler = () => {
        console.log('[Templates Modal] Closing modal');
        templatesModal.setAttribute('aria-hidden', 'true');
        templatesModal.style.display = 'none';
      };

      // Use addEventListener instead of onclick to ensure it's not overwritten
      templatesClose.addEventListener('click', (e) => {
        console.log('[Templates Modal] Close button clicked', e);
        console.log('[Templates Modal] Event target:', e.target);
        console.log('[Templates Modal] Current target:', e.currentTarget);
        e.preventDefault();
        e.stopPropagation();
        closeTemplatesModalHandler();
      }, true); // Use capture phase

      // Also add a direct onclick as fallback
      templatesClose.onclick = (e) => {
        console.log('[Templates Modal] Close button onclick fired');
        closeTemplatesModalHandler();
      };

      templatesModal.onclick = (e) => {
        if (e.target === templatesModal) {
          console.log('[Templates Modal] Background clicked');
          closeTemplatesModalHandler();
        }
      };
    } else {
      console.warn('[Templates Modal Setup] Some elements not found:', { templatesBtn, templatesModal, templatesClose });
    }

    setupAnnouncementModal(panel);

    setLanguage(currentLang);

    const multiToggle = panel.querySelector('#multiToggle');
    if (multiToggle) {
      multiToggle.checked = app.multiSlideMode;
      multiToggle.addEventListener('change', (e) => {
        app.multiSlideMode = e.target.checked;
        app.userSlideModeSet = true;
        renderPreview();
        updateHeaderButtons();
      });
    }

    if (!app.downloadIframe) {
      app.downloadIframe = document.createElement('iframe');
      app.downloadIframe.id = BOX_ID;
      app.downloadIframe.style.display = 'none';
      app.downloadIframe.setAttribute('sandbox', 'allow-scripts');
      app.downloadIframe.src = safeGetURL('sandbox/pptx-sandbox.html');
      app.downloadIframe.onload = () => { app.iframeReady = true; };
      document.body.appendChild(app.downloadIframe);
    }

    renderPreview();
    updateHeaderButtons();
    startObserveChanges();
  }

  function setupAnnouncementModal(panel) {
    if (!CURRENT_ANNOUNCEMENT || !CURRENT_ANNOUNCEMENT.id) return;
    const modal = panel.querySelector('#announcement-modal');
    if (!modal) return;

    updateAnnouncementLocale(panel, currentLang);

    const confirmBtn = panel.querySelector('#announcement-confirm');
    const closeBtn = panel.querySelector('#announcement-close');
    let persisted = app.announcementDismissed;

    const markDismissed = () => {
      if (persisted) return;
      persisted = true;
      const state = {
        id: CURRENT_ANNOUNCEMENT.id,
        dismissed: true,
        dismissedAt: new Date().toISOString(),
      };
      app.announcementState = state;
      app.announcementDismissed = true;
      saveAnnouncementState(state);
    };

    const closeModal = (persist) => {
      modal.classList.remove('show');
      modal.setAttribute('aria-hidden', 'true');
      if (persist) markDismissed();
    };

    const openModal = () => {
      modal.setAttribute('aria-hidden', 'false');
      modal.classList.add('show');
      setTimeout(() => {
        if (confirmBtn && typeof confirmBtn.focus === 'function') {
          confirmBtn.focus({ preventScroll: true });
        }
      }, 50);
    };

    const handleDismiss = () => closeModal(true);

    confirmBtn?.addEventListener('click', handleDismiss);
    closeBtn?.addEventListener('click', handleDismiss);
    modal.addEventListener('click', (event) => {
      if (event.target === modal) {
        handleDismiss();
      }
    });

    panel.addEventListener('keydown', (event) => {
      if (event.key === 'Escape' && modal.classList.contains('show')) {
        event.preventDefault();
        handleDismiss();
      }
    });

    if (!app.announcementDismissed) {
      requestAnimationFrame(() => {
        openModal();
      });
    }
  }

  function extractBlocks() {
      const nodes = [
        ...document.querySelectorAll('code.language-html, code.language-js, pre')
      ];
      if (!nodes.length) return { error: 'コードブロックが見当たりません' };

      let htmlText = '',
          pptxText = '',
          latest = '';

      for (let i = nodes.length - 1; i >= 0; i--) {
        const el = nodes[i];
        const txt = (el.innerText || el.textContent).trim();
        const isPptx = /(new\s+PptxGenJS|\.addSlide|\bPptxGenJS\b|\.slide\(|\.addText\(|\.addImage\(|\.addShape\()/.test(txt);
        const isHtml = !isPptx && /<\/?[a-zA-Z][^>]*>/.test(txt);

        if (!latest) {
          if (isPptx) latest = 'pptx';
          else if (isHtml) latest = 'html';
        }
        if (!htmlText && isHtml) {
          htmlText = txt;
          console.log('HTMLを検知');
        }
        if (!pptxText && isPptx) {
          pptxText = txt;
          console.log('pptxgenjsを検知');
        }
        if (htmlText && pptxText && latest) break;
      }

      if (!htmlText && !pptxText) {
        return { error: '対象コードがありません' };
      }
      return { html: htmlText, pptx: pptxText, latest };
    }

  function renderPreview() {
    const wrap = document.getElementById(app.PANEL_ID);
    if (!wrap) return;

    // Save current preview state before resetting
    const previousPreview = app.currentPreview;
    app.currentPreview = '';

    const jsonFound = hasSlidesJson();
    if (jsonFound && !app.userSlideModeSet) {
      if (!app.multiSlideMode) {
        app.multiSlideMode = true;
        const toggle = document.getElementById('multiToggle');
        if (toggle) toggle.checked = true;
        const radio = wrap.querySelector('#slide-mode-options input[value="multi"]');
        if (radio) radio.checked = true;
      }
    } else if (!jsonFound && app.multiSlideMode && (!app.htmlSlides || !app.htmlSlides.length) && !app.userSlideModeSet) {
      app.multiSlideMode = false;
      const toggle = document.getElementById('multiToggle');
      if (toggle) toggle.checked = false;
      const radio = wrap.querySelector('#slide-mode-options input[value="single"]');
      if (radio) radio.checked = true;
    }
    updateHeaderButtons();
    if (app.downloading) return;

    const view       = wrap.querySelector('#preview-content');
    const prevInput  = view.querySelector('#start-text');
    if (prevInput) {
      app.startText = prevInput.value;
    }
    const dlBtn      = wrap.querySelector('#download-btn');
    const copyBtn    = wrap.querySelector('#copy-html-btn');
    const convertBtn = wrap.querySelector('#convert-btn');
    let container  = view.querySelector('.preview-container');
    let allHtml    = app.multiSlideMode ? extractAllHtmlBlocks() : null;
    if (app.multiSlideMode && (!allHtml || !allHtml.length) && app.htmlSlides && app.htmlSlides.length) {
      allHtml = app.htmlSlides;
    }
    const { html: newHtml, pptx, latest } = extractBlocks();
    const currentPptxSnippets = extractAllPptxSnippets();
    if (newHtml) {
      app.lastHtml = newHtml;
      app.detectedHtml = true;
    } else if (pptx && latest === 'pptx') {
      app.lastHtml = '';
      app.detectedHtml = false;
    }
    if (pptx) {
      app.lastPptx = pptx;
      app.scrapedCode = pptx;
    } else if (!currentPptxSnippets.length) {
      app.lastPptx = '';
      if (!app.lastHtml) {
        app.scrapedCode = '';
      }
    }
    const slidesJson = hasSlidesJson() ? getSlidesJson() : null;
    let html = newHtml || app.lastHtml;
    if (app.multiSlideMode && allHtml && allHtml.length) {
      html = allHtml[allHtml.length - 1];
    } else if (!html && app.htmlSlides && app.htmlSlides.length) {
      html = app.htmlSlides[app.htmlSlides.length - 1];
    }
    if (app.multiSlideMode && allHtml && allHtml.length) {
      app.detectedHtml = true;
    } else if (!html) {
      app.detectedHtml = false;
    }
    const pptxCode = pptx || app.lastPptx;
    if ((html || pptxCode) && document.querySelector('#progress-indicator')) {
      hideProgress();
    }
    if (html || pptxCode) {
      hideWaitOverlay();
      if (!prevInput) resetSearchMode(true);
    }
    const diagramSrc = safeGetURL('images/Diagram.svg');
    const chartSrc   = safeGetURL('images/Chart.svg');


      function ensureProgressOverlay() {
        if (!view.querySelector('#progress-indicator')) {
          const div = document.createElement('div');
          div.id = 'progress-indicator';
          div.innerHTML = `
            <div class="pptx-progress-message" data-i18n="progressPreparingExport"></div>
            <div class="pptx-progress-bar"><div class="pptx-progress"></div></div>
            <div class="pptx-progress-text">0%</div>
            <button type="button" class="pptx-cancel" data-i18n="cancel"></button>
          `;
          view.appendChild(div);
          applyTranslations(div);
        }
      }

      function ensureWaitOverlay() {
        if (!view.querySelector('#wait-overlay')) {
          const div = document.createElement('div');
          div.id = 'wait-overlay';
          div.innerHTML = `
            <div class="wait-spinner"></div>
            <div class="wait-message" data-i18n="pleaseWait"></div>
          `;
          view.appendChild(div);
          applyTranslations(div);
        }
      }


      function ensureArrowOverlay() {
        if (!view.querySelector('#arrow-overlay')) {
          const div = document.createElement('div');
          div.id = 'arrow-overlay';
          view.appendChild(div);
        }
      }

      function ensurePptxOverlay() {
        if (!view.querySelector('#pptx-detected-overlay')) {
          const div = document.createElement('div');
          div.id = 'pptx-detected-overlay';
          div.innerHTML = `
            <div data-i18n="exportReady"></div>
            <div class=\"pptx-overlay-actions\">
              <button type=\"button\" id=\"pptx-only-download\" data-i18n=\"exportPptx\"></button>
              <button type=\"button\" id=\"pptx-fast-download\" data-i18n=\"downloadViaApi\"></button>
            </div>`;
          view.appendChild(div);
          applyTranslations(div);
        }
      }

      function ensurePreviewContainer() {
        if (!container || !view.contains(container)) {
          view.innerHTML = '';
          container = document.createElement('div');
          container.className = 'preview-container';
          view.appendChild(container);
        }
        container.style.width = '100%';
        container.style.height = '';
      }

    const homeBtn = wrap.querySelector('#home-btn');
    if (homeBtn) {
      homeBtn.classList.toggle('active', app.showHome);
      const img = homeBtn.querySelector('img');
      if (img) img.src = safeGetURL(app.showHome ? app.HOME_WHITE_SRC : app.HOME_GRAY_SRC);
    }

    if (app.showHome || (!html && !pptxCode && !slidesJson)) {
      app.currentPreview = 'home';
      const registered = app.previewRegDone && app.pptxRegDone;
      const regHtml = `<div id="reg-info" style="margin:0 20px 10px;text-align:center;" data-i18n="regInfo"></div>
        <div id="reg-buttons" style="margin-bottom:10px;">
          <button type="button" id="preview-reg-btn" style="margin-right:5px;" data-i18n="registerPreviewAi"></button>
          <button type="button" id="pptx-reg-btn" data-i18n="registerPptAi"></button>
        </div>`;
      if (!registered) {
        view.innerHTML = `<div style="width:100%;height:100%;display:flex;flex-direction:column;justify-content:center;align-items:center;">
          ${regHtml}
        </div>`;
      } else {
        view.innerHTML = `<div style="width:100%;height:100%;display:flex;flex-direction:column;justify-content:center;align-items:center;">
          <div data-i18n="whatSlide"></div>
          <div id="slide-type-options">
            <div class="option">
              <img src="${diagramSrc}" data-prompt="下記の内容をもとにわかりやすい図解のスライドを作成して。">
              <div class="label-overlay" data-i18n="diagram"></div>
            </div>
            <div class="option">
              <img src="${chartSrc}" data-prompt="下記の内容をもとにグラフを活用したわかりやすいスライドを作成して。">
              <div class="label-overlay" data-i18n="graph"></div>
            </div>
          </div>
          <div id="slide-mode-options">
            <label class="slide-mode-option">
              <input type="radio" name="slide-mode" value="single" checked>
              <div class="mode-text">
                <span class="mode-title" data-i18n="singleSlide"></span>
                <span class="mode-desc" data-i18n="editableShort"></span>
              </div>
            </label>
            <label class="slide-mode-option">
              <input type="radio" name="slide-mode" value="multi">
              <div class="mode-text">
                <span class="mode-title" data-i18n="multiSlideMode"></span>
                <span class="mode-desc" data-i18n="notEditableLong"></span>
              </div>
              <span class="beta-tag">Beta</span>
            </label>
          </div>
          <div style="width:80%;margin:10px auto 0;display:flex;gap:8px;">
            <button type="button" id="pptx-upload-text-btn" style="flex:1;padding:8px 16px;background:#f8f9fa;border:1px solid #ddd;border-radius:4px;cursor:pointer;font-size:13px;color:#555;margin-bottom:8px;" data-i18n="analyzePptx"></button>
            <button type="button" id="home-templates-btn" style="flex:1;padding:8px 16px;background:#f8f9fa;border:1px solid #ddd;border-radius:4px;cursor:pointer;font-size:13px;color:#555;margin-bottom:8px;" data-i18n="templatesTitle"></button>
          </div>
          <textarea id="start-text" style="width:80%;height:240px;margin-top:0;" data-i18n-placeholder="placeholderStartText"></textarea>
          <div id="prompt-length-warning" data-i18n="promptTooLong"></div>
          <div style="margin-top:10px;display:flex;align-items:center;justify-content:center;">
            <button type="button" id="start-chat-btn" disabled data-i18n="create"></button>
            <div class="toggle-wrap"><span class="toggle-label">Search:</span><select id="search-mode-select"><option value="off">Off</option><option value="web">Web search</option></select></div>
          </div>
        </div>`;
      }
      applyTranslations(view);
      ensureProgressOverlay();
      ensureWaitOverlay();
      ensurePptxOverlay();
      ensureArrowOverlay();
      dlBtn.disabled = true;
      dlBtn.classList.remove('active');
      if (convertBtn) {
        convertBtn.disabled = true;
        convertBtn.classList.remove('active');
        updateConvertBtn(convertBtn, 'convert');
      }
      if (copyBtn) {
        copyBtn.style.display = 'none';
        copyBtn.disabled = true;
        copyBtn.classList.remove('active');
        copyBtn.textContent = 'Copy HTML';
      }
      const startBtn = view.querySelector('#start-chat-btn');
      const modeSel  = view.querySelector('#search-mode-select');
      const startBox = view.querySelector('#start-text');
      if (startBox) {
        startBox.value = app.startText;
        startBox.onfocus = () => { app.isTyping = true; };
        startBox.onblur = () => {
          app.isTyping = false;
          if (app.pendingRender) {
            app.pendingRender = false;
            renderPreview();
          }
        };
      }
      const slides   = view.querySelectorAll('#slide-type-options img');
      if (slides && slides.length) {
        slides.forEach(img => {
          if (app.slidePrompt && img.dataset.prompt === app.slidePrompt) {
            img.classList.add('selected');
          }
          img.onclick = () => {
            const isSelected = img.classList.contains('selected');
            slides.forEach(i => i.classList.remove('selected'));
            if (isSelected) {
              app.slidePrompt = '';
            } else {
              img.classList.add('selected');
              app.slidePrompt = img.dataset.prompt || '';
            }
          };
        });
      }
      const slideModeRadios = view.querySelectorAll('#slide-mode-options input');
      if (slideModeRadios && slideModeRadios.length) {
        slideModeRadios.forEach(radio => {
          radio.checked = app.multiSlideMode ? radio.value === 'multi' : radio.value === 'single';
          radio.onchange = () => {
            app.multiSlideMode = radio.value === 'multi';
            app.userSlideModeSet = true;
            const toggle = document.getElementById('multiToggle');
            if (toggle) toggle.checked = app.multiSlideMode;
            renderPreview();
            updateHeaderButtons();
          };
        });
      }

      if (modeSel) {
        if (app.searchMode === 'deep') {
          app.searchMode = 'off';
        }
        modeSel.value = app.searchMode;
        modeSel.onchange = () => {
          app.searchMode = modeSel.value;
          saveRegStatus(app);
        };
      }
      if (startBtn && startBox) {
        const hasText = startBox.value.trim().length > 0;
        startBtn.disabled = !hasText;
        startBtn.classList.toggle('active', hasText);
        const warnEl = view.querySelector('#prompt-length-warning');
        const updateWarning = () => {
          const has = startBox.value.trim().length > 0;
          app.startText = startBox.value;
          startBtn.disabled = !has;
          startBtn.classList.toggle('active', has);
          const len = startBox.value.length;
          if (warnEl) warnEl.style.display = len >= 100000 ? 'block' : 'none';
        };
        startBox.oninput = updateWarning;
        updateWarning();
        startBtn.onclick = () => {
          const txt = (startBox.value || '').trim();
          app.startText = startBox.value;
          let finalText = txt;
          if (app.multiSlideMode && txt) {
            finalText = MULTI_SLIDE_PREFIX + '\n\n' + txt;
          }
          let promptText = '';
          if (app.searchMode === 'deep') {
            promptText = finalText;
          } else {
            const head = app.slidePrompt ? app.slidePrompt + ' ' : '';
            promptText = head + finalText;
          }
          startChat(promptText, app.slidePrompt);
        };
      }
      const previewReg = view.querySelector('#preview-reg-btn');
      if (previewReg) {
        if (app.previewRegDone) {
          markRegButtonDone(previewReg);
        }
        previewReg.onclick = () => {
          openTab(app.START_CHAT_URL);
          // タブが開かれるのを待ってから登録完了状態にする
          setTimeout(() => {
            app.previewRegDone = true;
            saveRegStatus(app);
            markRegButtonDone(previewReg);
            scheduleHideRegElements(true);
          }, 500);
        };
      }
      const pptxReg = view.querySelector('#pptx-reg-btn');
      if (pptxReg) {
        if (app.pptxRegDone) {
          markRegButtonDone(pptxReg);
        }
        pptxReg.onclick = () => {
          openTab(app.CHAT_URL);
          // タブが開かれるのを待ってから登録完了状態にする
          setTimeout(() => {
            app.pptxRegDone = true;
            saveRegStatus(app);
            markRegButtonDone(pptxReg);
            scheduleHideRegElements(true);
          }, 500);
        };
      }

      // Setup PPTX upload button in home screen
      const pptxUploadTextBtn = view.querySelector('#pptx-upload-text-btn');
      if (pptxUploadTextBtn && app.openPptxModal) {
        pptxUploadTextBtn.onclick = app.openPptxModal;
      }

      // Setup Templates button in home screen
      const homeTemplatesBtn = view.querySelector('#home-templates-btn');
      if (homeTemplatesBtn) {
        homeTemplatesBtn.onclick = () => {
          console.log('[Home] Templates button clicked');
          openTemplatesModal();
        };
      }

      scheduleHideRegElements(false);
      return;
    }

      ensurePreviewContainer();
      container.classList.toggle('multi-slide', app.multiSlideMode);
      if (app.multiSlideMode) {
        if (allHtml && allHtml.length) {
          app.currentPreview = 'html';
          container.innerHTML = '';
          ensureProgressOverlay();
          ensureWaitOverlay();
          ensurePptxOverlay();
          let w = container.clientWidth || wrap.clientWidth;
          const h = w * 720 / 1280;
          allHtml.forEach(htmlBlock => {
          const slide = document.createElement('div');
          slide.className = 'preview-slide';
          Object.assign(slide.style, { position: 'relative', width: `${w}px`, height: `${h}px` });
          const frame = document.createElement('iframe');
          frame.setAttribute('sandbox', 'allow-scripts');
          frame.className = 'slide-frame';
          frame.srcdoc = resolvePreviewAssets(htmlBlock);
          Object.assign(frame.style, {
            position: 'absolute',
            top: 0,
            left: 0,
            width: '1280px',
            height: '720px',
            transformOrigin: 'top left',
            transform: `scale(${w / 1280})`,
          });
          frame.setAttribute('scrolling', 'no');
          frame.style.overflow = 'hidden';
            slide.appendChild(frame);
            container.appendChild(slide);
          });
          if (copyBtn) {
            copyBtn.style.display = 'inline-block';
            copyBtn.disabled = false;
            copyBtn.classList.add('active');
            copyBtn.textContent = 'Copy HTML';
          }
          if (convertBtn) {
            convertBtn.disabled = false;
            convertBtn.classList.add('active');
            updateConvertBtn(convertBtn, 'convert');
          }
        } else if (slidesJson && slidesJson.length) {
          app.currentPreview = 'json';
          container.innerHTML = '';
          ensureProgressOverlay();
          ensureWaitOverlay();
          ensurePptxOverlay();
          let w = container.clientWidth || wrap.clientWidth;
          const h = w * 720 / 1280;
          slidesJson.forEach(sl => {
            const slide = document.createElement('div');
            slide.className = 'preview-slide';
            Object.assign(slide.style, { position: 'relative', width: `${w}px`, height: `${h}px` });
            const inner = document.createElement('div');
            inner.className = 'json-slide';
            Object.assign(inner.style, {
              position: 'absolute',
              top: 0,
              left: 0,
              width: '1280px',
              height: '720px',
              transformOrigin: 'top left',
              transform: `scale(${w / 1280})`,
              display: 'flex',
              flexDirection: 'column',
              justifyContent: 'flex-start',
              alignItems: 'flex-start',
              boxSizing: 'border-box',
              padding: '40px',
              textAlign: 'left'
            });
            const title = document.createElement('h1');
            title.className = 'json-slide-title';
            title.textContent = sl.title || '';
            const txt = document.createElement('div');
            txt.className = 'json-slide-prompt';
            txt.textContent = (sl.prompt || '').replace(/\n/g, '\n');
            inner.appendChild(title);
            inner.appendChild(txt);
            slide.appendChild(inner);
            container.appendChild(slide);
          });
          if (copyBtn) {
            copyBtn.style.display = 'none';
            copyBtn.disabled = true;
            copyBtn.classList.remove('active');
          }
          if (convertBtn) {
            convertBtn.disabled = true;
            convertBtn.classList.remove('active');
            updateConvertBtn(convertBtn, 'convert');
          }
        }
      } else if (html && (newHtml || !container.querySelector('iframe') || html !== app.lastHtml)) {
        app.currentPreview = 'html';
        if (newHtml) app.lastHtml = newHtml;
        container.innerHTML = '';
        ensureProgressOverlay();
        ensureWaitOverlay();
        ensurePptxOverlay();
        let w = container.clientWidth || wrap.clientWidth;
        const h = w * 720 / 1280;
        Object.assign(container.style, { width: `${w}px`, height: `${h}px` });

        const frame = document.createElement('iframe');
        frame.setAttribute('sandbox', 'allow-scripts');
        frame.className = 'slide-frame';
        frame.srcdoc = resolvePreviewAssets(html);
        Object.assign(frame.style, {
          position: 'absolute',
          top: 0,
          left: 0,
          width: '1280px',
          height: '720px',
          transformOrigin: 'top left',
          transform: `scale(${w / 1280})`,
        });
        frame.setAttribute('scrolling', 'no');
        frame.style.overflow = 'hidden';
        container.appendChild(frame);
        if (copyBtn) {
          copyBtn.style.display = 'inline-block';
          copyBtn.disabled = false;
          copyBtn.classList.add('active');
          copyBtn.textContent = 'Copy HTML';
        }
        if (convertBtn) {
          convertBtn.disabled = false;
          convertBtn.classList.add('active');
          updateConvertBtn(convertBtn, 'convert');
        }
      }

      const isPptxLatest = latest === 'pptx' && pptxCode;

      let overlay = view.querySelector('#pptx-detected-overlay');

      const bindPptxButtons = () => {
        const overlayEl = view.querySelector('#pptx-detected-overlay');
        if (overlayEl) {
          const inDlBtn = overlayEl.querySelector('#pptx-only-download');
          if (inDlBtn) inDlBtn.onclick = () => app.downloadPptx && app.downloadPptx();

          // 高速ダウンロードボタン
          const fastDownloadBtn = overlayEl.querySelector('#pptx-fast-download');
          if (fastDownloadBtn) {
            fastDownloadBtn.onclick = async () => {
              try {
                if (!app.scrapedCode) {
                  alert(t('noTargetCode'));
                  return;
                }

                // APIクライアントを動的にインポート
                const { generatePptxViaApi, getApiKey } = await import(chrome.runtime.getURL('src/apiClient.js'));

                // APIキーの確認
                const apiKey = await getApiKey();
                if (!apiKey) {
                  alert(t('apiKeyNotSet'));
                  return;
                }

                // 進捗表示を開始
                app.showProgress(t('apiGenerating'), 120000);

                // ファイル名を生成
                let fileName = 'presentation.pptx';
                try {
                  const m = app.scrapedCode.match(/addText\(\s*(["'`])([\s\S]*?)\1/);
                  if (m) {
                    const text = m[2].replace(/\\n/g, ' ').trim();
                    const sanitized = text.replace(/[\\/:*?"<>|]/g, '_').slice(0, 20);
                    const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
                    fileName = `${sanitized}_${date}.pptx`;
                  }
                } catch (e) {
                  console.error('[API Download] ファイル名生成エラー', e);
                }

                // API経由でパワーポイントを生成
                await generatePptxViaApi(app.scrapedCode, fileName);

                // 進捗表示を終了
                app.stopFakeProgress();
                app.updateProgress(100);
                setTimeout(() => {
                  app.hideProgress();
                  // alert(t('apiDownloadSuccess')); // アラート削除：ダウンロード成功は通知不要
                }, 500);

              } catch (error) {
                console.error('[API Download] エラー:', error);
                app.stopFakeProgress();
                app.hideProgress();

                // エラーメッセージの判定
                let errorMsg = t('apiDownloadFailed');
                if (error.message.includes('APIキーが設定されていません')) {
                  errorMsg = t('apiKeyNotSet');
                } else if (error.message.includes('APIキーが無効')) {
                  errorMsg = t('apiKeyInvalid');
                } else {
                  errorMsg = `${t('apiDownloadFailed')}: ${error.message}`;
                }

                alert(errorMsg);
              }
            };
          }
        }
        const csvTemplateBtn = wrap.querySelector('#pptx-download-csv');
        if (csvTemplateBtn) csvTemplateBtn.onclick = () => downloadCsvTemplateFromPptx();
      };

      // Check if preview is already being displayed
      // Use previousPreview since app.currentPreview was reset at the beginning
      // Check for actual preview content (not just overlays)
      const hasPreviewContent = container && (
        container.querySelector('.pptx-preview-wrapper') || // PPTX preview
        container.querySelector('iframe') || // HTML preview
        (container.children.length > 0 && !container.querySelector('#pptx-detected-overlay')) // Other content
      );
      const hasPptxPreview = hasPreviewContent && previousPreview === 'pptx';
      const hasHtmlPreview = hasPreviewContent && previousPreview === 'html';

      console.log('[Preview] hasPptxPreview:', hasPptxPreview, 'hasHtmlPreview:', hasHtmlPreview, 'previousPreview:', previousPreview);

      // Logic:
      // 1. PPTX preview + new HTML detected → switch to HTML (clear currentPreview)
      // 2. PPTX preview + PPTX detected → keep preview (restore currentPreview)
      // 3. HTML preview + PPTX detected → show overlay (don't keep, clear currentPreview)

      if (newHtml && hasPptxPreview) {
        // Case 1: New HTML detected while showing PPTX preview - switch to HTML
        console.log('[Preview] New HTML detected, switching from PPTX to HTML');
        app.currentPreview = '';
      } else if (hasPptxPreview) {
        // Case 2: Keep PPTX preview
        console.log('[Preview] Keeping PPTX preview');
        app.currentPreview = previousPreview;
      }
      // Case 3: HTML preview - don't keep, let overlay show (no action needed, currentPreview stays '')

      // Only keep PPTX preview, not HTML preview
      const shouldKeepPreview = hasPptxPreview && !newHtml;

      if (html && isPptxLatest && !shouldKeepPreview) {
        // Show overlay when PPTX is detected (unless keeping PPTX preview)
        ensurePptxOverlay();
        overlay = view.querySelector('#pptx-detected-overlay');
        if (overlay) {
          overlay.style.display = 'flex';
        }
      } else if (!html && isPptxLatest && !shouldKeepPreview) {
        // No HTML, only PPTX, and not keeping preview
        app.currentPreview = 'pptx';
        view.innerHTML = '';
        ensureProgressOverlay();
        ensureWaitOverlay();
        ensurePptxOverlay();
        overlay = view.querySelector('#pptx-detected-overlay');
        if (overlay) {
          overlay.style.display = 'flex';
        }
      } else if (overlay && !shouldKeepPreview) {
        overlay.style.display = 'none';
      } else if (overlay && shouldKeepPreview) {
        // Keep overlay hidden if keeping PPTX preview
        overlay.style.display = 'none';
      }

      bindPptxButtons();

      const modalCsvBtn = wrap.querySelector('#pptx-download-csv');
      if (modalCsvBtn) {
        const hasPptxCode = !!pptxCode;
        const shouldShowCsv = (!html || isPptxLatest) || !hasPptxCode;
        modalCsvBtn.style.display = shouldShowCsv ? 'inline-flex' : 'none';
        modalCsvBtn.disabled = !shouldShowCsv;
      }


        if (isPptxLatest && pptxCode) {
          dlBtn.disabled = false;
          dlBtn.classList.add('active');
          if (convertBtn) {
            convertBtn.disabled = true;
            convertBtn.classList.remove('active');
            updateConvertBtn(convertBtn, 'converted');
          }
        } else if (html) {
          dlBtn.disabled = true;
          dlBtn.classList.remove('active');
          if (convertBtn) {
            convertBtn.disabled = false;
            convertBtn.classList.add('active');
            updateConvertBtn(convertBtn, 'convert');
          }
        } else if (pptxCode) {
          app.currentPreview = 'pptx';
          dlBtn.disabled = false;
          dlBtn.classList.add('active');
          if (convertBtn) {
            convertBtn.disabled = true;
            convertBtn.classList.remove('active');
            updateConvertBtn(convertBtn, 'converted');
          }
        } else {
          app.scrapedCode = '';
          dlBtn.disabled = true;
          dlBtn.classList.remove('active');
        if (convertBtn) {
          convertBtn.disabled = true;
          convertBtn.classList.remove('active');
          updateConvertBtn(convertBtn, 'convert');
        }
      }

    fitPreviewSlides();
    adjustPreviewHeight();
    updateActionStates(
      app.multiSlideMode ? (allHtml && allHtml.length) : !!html,
      extractAllPptxSnippets().length > 0,
      Array.isArray(slidesJson) && slidesJson.length > 0
    );
    updateHeaderButtons();

    updatePromptCandidates();

    const dlCancel = view.querySelector('#progress-indicator .pptx-cancel');
    if (dlCancel) dlCancel.onclick = cancelDownload;
  }

  // Return an array of HTML code strings for all html code blocks on the page
  function extractAllHtmlBlocks() {
    try {
      return [...document.querySelectorAll('code.language-html')]
        .map(el => (el.textContent || el.innerText || '').trim())
        .filter(Boolean);
    } catch (e) {
      console.error('extractAllHtmlBlocks error', e);
      return [];
    }
  }

  // ページ内のすべてのPptxGenJSコードブロックを集める
  function extractAllPptxSnippets() {
    try {
      return [...document.querySelectorAll('code.language-js, code.language-javascript')]
        .map(el => (el.textContent || el.innerText || '').trim())
        .filter(txt => /(new\s+PptxGenJS|\.addSlide|\bPptxGenJS\b|\.addText\(|\.addImage\(|\.addShape\()/
          .test(txt));
    } catch (e) {
      console.error('extractAllPptxSnippets error', e);
      return [];
    }
  }

  // 指定数のPPTXコードが見つかるまで待機する
  async function waitForPptxSnippets(expected, timeout = 10000) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
      if (extractAllPptxSnippets().length >= expected) return true;
      await sleep(200);
    }
    return false;
  }

  // コード内のテキストからファイル名を作成する
  function buildFileNameFromSnippet(snippet) {
    try {
      const m = snippet.match(/addText\(\s*(["'`])([\s\S]*?)\1/);
      if (m) {
        const text = m[2].replace(/\\n/g, ' ').trim();
        const sanitized = text.replace(/[\\/:*?"<>|]/g, '_').slice(0, 20);
        const date = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        return `${sanitized}_${date}.pptx`;
      }
    } catch (e) {
      console.error('buildFileNameFromSnippet error', e);
    }
    return 'slides.pptx';
  }

  // サンドボックス用のiframeを確実に生成する
  async function ensureSandboxIframe() {
    if (!app.downloadIframe) {
      app.downloadIframe = document.createElement('iframe');
      app.downloadIframe.id = app.BOX_ID;
      app.downloadIframe.style.display = 'none';
      app.downloadIframe.setAttribute('sandbox', 'allow-scripts');
      app.downloadIframe.src = safeGetURL('sandbox/pptx-sandbox.html');
      document.body.appendChild(app.downloadIframe);
    }
    if (app.iframeReady) return app.downloadIframe;
    return await new Promise(resolve => {
      app.downloadIframe.onload = () => {
        app.iframeReady = true;
        resolve(app.downloadIframe);
      };
    });
  }

  // 複数のコードからまとめてスライドを生成する
  async function generateMultipleSlides(codes) {
    await ensureSandboxIframe();
    const snippets = Array.isArray(codes) ? codes : [];
    return await new Promise((resolve, reject) => {
      const iframe = app.downloadIframe;
      const iframeOrigin = new URL(iframe.src, window.location.href).origin;
      const sandboxed = iframe.sandbox && !iframe.sandbox.contains('allow-same-origin');
      const target = sandboxed ? '*' : iframeOrigin;
      const handle = (e) => {
        if (!e || !e.data || e.source !== iframe.contentWindow) return;
        if (!sandboxed && e.origin !== iframeOrigin) return;
        if (sandboxed && e.origin !== 'null') return;
        if (e.data.action === 'multi-generated') {
          window.removeEventListener('message', handle);
          resolve(e.data.blob);
        } else if (e.data.action === 'error') {
          window.removeEventListener('message', handle);
          reject(new Error(e.data.message));
        }
      };
      window.addEventListener('message', handle);
      try {
        iframe.contentWindow.postMessage({ action: 'generate-multi', codes: snippets }, target);
      } catch (err) {
        window.removeEventListener('message', handle);
        reject(err);
      }
    });
  }

  // 収集したコードをまとめてPPTXとしてダウンロード
  async function downloadMultiplePptx() {
    const snippets = extractAllPptxSnippets();
    if (!snippets.length) return;
    const fileName = buildFileNameFromSnippet(snippets[0]);
    // ローディング時間は6分に設定
    showProgress('exportingPptx', 360000);
    try {
      const blob = await generateMultipleSlides(snippets);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      updateProgress(100);
      updateProgressMessage('exportDone');
    } catch (e) {
      console.error('download multiple pptx failed', e);
      updateProgressMessage('exportFailed');
      alert(t('exportFailed'));
      setTimeout(() => insertErrorPrompt(e.message), 700);
    }
    hideProgress();
  }

  // 進捗バーに表示するメッセージを変更
  function updateProgressMessage(key) {
    const el = document.querySelector(`#${app.PANEL_ID} .pptx-progress-message`);
    if (el && key) {
      el.dataset.i18n = key;
      el.textContent = t(key);
    }
  }

  let hideProgressTimer;

  // 進捗バーを表示してダウンロード中と示す
  function showProgress(key, duration, markDownloading = true) {
    if (hideProgressTimer) {
      clearTimeout(hideProgressTimer);
      hideProgressTimer = null;
    }
    if (!document.getElementById(app.PANEL_ID)) {
      openPanel();
    }
    const ind   = document.querySelector(`#${app.PANEL_ID} #progress-indicator`);
    const bar   = document.querySelector(`#${app.PANEL_ID} .pptx-progress`);
    const text  = document.querySelector(`#${app.PANEL_ID} .pptx-progress-text`);
    const pptxOverlay = document.querySelector(`#${app.PANEL_ID} #pptx-detected-overlay`);
    hideWaitOverlay();
    resetSearchMode(true);
    app.progressValue = 0;
    if (bar) bar.style.width = '0%';
    if (text) text.textContent = '0%';
    if (ind) ind.style.display = 'flex';
    if (pptxOverlay) pptxOverlay.style.display = 'none';
    updateProgressMessage(key || 'progressPreparingExport');
    startFakeProgress(duration);
    if (markDownloading) app.downloading = true;
  }

  // 進捗バーを非表示にする
  function hideProgress() {
    stopFakeProgress();
    app.progressValue = 100;
    updateProgress(app.progressValue);
    const ind = document.querySelector(`#${app.PANEL_ID} #progress-indicator`);
    if (hideProgressTimer) {
      clearTimeout(hideProgressTimer);
      hideProgressTimer = null;
    }
    if (ind) {
      hideProgressTimer = setTimeout(() => {
        ind.style.display = 'none';
        hideProgressTimer = null;
      }, 500);
    }
    hideWaitOverlay();
    app.downloading = false;
  }

  // 進捗バーの数値と幅を更新
  function updateProgress(val) {
    const bar  = document.querySelector(`#${app.PANEL_ID} .pptx-progress`);
    const text = document.querySelector(`#${app.PANEL_ID} .pptx-progress-text`);
    app.progressValue = val;
    if (bar) bar.style.width = app.progressValue + '%';
    if (text) text.textContent = app.progressValue + '%';
  }


  // 実際の進捗がなくても少しずつ進む演出を開始
  function startFakeProgress(duration = 120000) {
    stopFakeProgress();
    const start = Date.now();
    app.fakeProgressTimer = setInterval(() => {
      const elapsed = Date.now() - start;
      const ratio = Math.min(elapsed / duration, 0.99);
      const val = Math.floor(ratio * 100);
      if (val > app.progressValue) updateProgress(val);
      if (ratio >= 0.99) stopFakeProgress();
    }, 1000);
  }

  // 擬似進捗のタイマーを止める
  function stopFakeProgress() {
    if (app.fakeProgressTimer) {
      clearInterval(app.fakeProgressTimer);
      app.fakeProgressTimer = null;
    }
  }

  // プレビュー関連のボタンを一時的に無効化
  function disablePreviewButtons() {
    const wrap = document.getElementById(app.PANEL_ID);
    if (!wrap) return;
    const ids = ['convert-btn', 'download-btn', 'copy-html-btn'];
    app.prevDisabled = {};
    ids.forEach(id => {
      const btn = wrap.querySelector('#' + id);
      if (btn) {
        app.prevDisabled[id] = btn.disabled;
        btn.disabled = true;
      }
    });
  }

  // 無効化したボタンを元の状態に戻す
  function restorePreviewButtons() {
    if (!app.prevDisabled) return;
    const wrap = document.getElementById(app.PANEL_ID);
    if (!wrap) return;
    Object.entries(app.prevDisabled).forEach(([id, state]) => {
      const btn = wrap.querySelector('#' + id);
      if (btn) btn.disabled = state;
    });
    app.prevDisabled = null;
  }


  // 次の処理が始まるまで待機中であることを示す
  function showWaitOverlay() {
    if (!document.getElementById(app.PANEL_ID)) {
      openPanel();
    }
    const overlay = document.querySelector(`#${app.PANEL_ID} #wait-overlay`);
    if (overlay) overlay.style.display = 'flex';
    const params = new URLSearchParams(location.search);
    const mode = params.get(app.SEARCH_PARAM) || (params.get(app.DEEP_PARAM) === '1' ? 'deep' : '');
    if (!mode) resetSearchMode(true);
  }

  // 待機オーバーレイを閉じる
  function hideWaitOverlay() {
    const overlay = document.querySelector(`#${app.PANEL_ID} #wait-overlay`);
    if (overlay) overlay.style.display = 'none';
  }

  // ======================
  // Template Masking & Save Functions
  // ======================

  // マスキング処理関数
  function performMasking(code, options) {
    const { maskText, maskBullet, maskChart, randomValues, preserveSpaces } = options;

    function maskString(text) {
      if (preserveSpaces) {
        return text.split('').map(char =>
          (char === ' ' || char === '\n' || char === '\t') ? char : 'X'
        ).join('');
      }
      return 'X'.repeat(text.length);
    }

    let maskedCode = code;

    // 通常テキストのマスキング
    if (maskText) {
      maskedCode = maskedCode.replace(/slide\.addText\((["'])([^"']*?)\1/g,
        (m, q, t) => `slide.addText(${q}${maskString(t)}${q}`);
      maskedCode = maskedCode.replace(/slide\.addText\('([^']+)'/g,
        (m, t) => `slide.addText('${maskString(t)}'`);
      maskedCode = maskedCode.replace(/slide\.addText\("([^"]+)"/g,
        (m, t) => `slide.addText("${maskString(t)}"`);
    }

    // 箇条書きテキストのマスキング
    if (maskBullet) {
      maskedCode = maskedCode.replace(/text:\s*(["'])([^"']*?)\1/g,
        (m, q, t, offset) => {
          const before = maskedCode.substring(Math.max(0, offset - 50), offset);
          if (before.includes('addText') || before.includes('[') || before.includes('{')) {
            return `text: ${q}${maskString(t)}${q}`;
          }
          return m;
        });
    }

    // グラフ名のマスキング
    if (maskChart) {
      maskedCode = maskedCode.replace(/(\baddChart[^}]*?)name:\s*(["'])([^"']*?)\2/g,
        (m, prefix, q, t) => `${prefix}name: ${q}${maskString(t)}${q}`);
    }

    // グラフ値のランダム化
    if (randomValues) {
      maskedCode = maskedCode.replace(/values:\s*\[([^\]]+)\]/g, (m, vals) => {
        const randomVals = vals.split(',').map(v => {
          const num = parseInt(v.trim());
          if (!isNaN(num)) {
            const min = Math.floor(num * 0.7);
            const max = Math.floor(num * 1.3);
            return Math.floor(Math.random() * (max - min + 1)) + min;
          }
          return v;
        }).join(', ');
        return `values: [${randomVals}]`;
      });
    }

    return maskedCode;
  }

  /**
   * マスキングしたPptxgenjsコードからプレビューHTMLを生成
   * @param {string} maskedCode - マスキング済みのPptxgenjsコード
   * @returns {Promise<string>} プレビューHTML
   */
  async function generatePreviewFromMaskedCode(maskedCode) {
    console.log('[Template Preview] マスキング版プレビュー生成開始');

    try {
      // 1. APIクライアントをインポート
      const { getApiKey } = await import(chrome.runtime.getURL('src/apiClient.js'));

      // 2. APIキーの確認
      const apiKey = await getApiKey();
      if (!apiKey) {
        throw new Error('APIキーが設定されていません');
      }

      // 3. 一時ファイル名を生成
      const tempFileName = `temp_preview_${Date.now()}.pptx`;

      // 4. マスキングしたコードをAPI経由でPPTX化
      console.log('[Template Preview] API経由でPPTX生成中...');

      const requestBody = {
        script: maskedCode,
        filename: tempFileName,
        payload: {}
      };

      // Background service worker経由でAPI呼び出し
      const responseData = await new Promise((resolve, reject) => {
        chrome.runtime.sendMessage({
          action: 'api-fetch',
          url: 'https://powerpoint-genai-test-854259963531.asia-northeast1.run.app/generate-pptx',
          options: {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'X-API-Key': apiKey,
              'X-Extension-ID': chrome.runtime.id
            },
            body: JSON.stringify(requestBody)
          }
        }, (response) => {
          if (chrome.runtime.lastError) {
            reject(new Error(chrome.runtime.lastError.message));
            return;
          }
          if (!response || !response.success) {
            reject(new Error(response?.error?.message || 'API呼び出しに失敗'));
            return;
          }
          resolve(response);
        });
      });

      const apiResponse = responseData.response;

      if (apiResponse.status !== 200) {
        throw new Error(`API Error: ${apiResponse.status}`);
      }

      // 5. レスポンスからBase64データを取得
      const data = (typeof apiResponse.data === 'object') ? apiResponse.data : JSON.parse(apiResponse.data);

      if (!data.data) {
        throw new Error('レスポンスにデータが含まれていません');
      }

      console.log('[Template Preview] Base64データ取得成功');

      // 6. Base64からArrayBufferに変換
      const base64Data = data.data;
      const binaryString = atob(base64Data);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      const arrayBuffer = bytes.buffer;

      console.log('[Template Preview] PPTX生成完了、プレビュー変換開始');

      // 7. pptx-previewライブラリを使ってHTMLプレビューを生成
      // グローバルスコープから取得（content_scriptsで読み込み済み）
      if (typeof pptxPreview === 'undefined' || typeof pptxPreview.init !== 'function') {
        throw new Error('pptx-preview library is not available');
      }
      const lib = pptxPreview;

      // 8. 一時的なコンテナを作成してプレビューを生成
      const tempContainer = document.createElement('div');
      tempContainer.style.cssText = 'position: absolute; left: -9999px; top: -9999px; width: 720px;';
      document.body.appendChild(tempContainer);

      const tempWrapper = document.createElement('div');
      tempContainer.appendChild(tempWrapper);

      const viewer = lib.init(tempWrapper, {
        width: 720,
        height: 405
      });

      await viewer.preview(arrayBuffer);

      // 9. プレースホルダーを削除
      setTimeout(() => {
        const placeholders = tempContainer.querySelectorAll('.pptx-preview-loading, .pptx-preview-slide-num');
        placeholders.forEach(el => el.remove());
      }, 100);

      // 少し待機してDOMが更新されるのを待つ
      await new Promise(resolve => setTimeout(resolve, 200));

      // 10. HTMLを取得して後処理
      let previewHtml = tempContainer.innerHTML;

      // 背景色を白に変換
      previewHtml = previewHtml.replace(/background:\s*rgb\(0,\s*0,\s*0\)/gi, 'background: rgb(255, 255, 255)');
      previewHtml = previewHtml.replace(/background:\s*#000000/gi, 'background: #FFFFFF');
      previewHtml = previewHtml.replace(/background:\s*black/gi, 'background: white');

      // マージンを削除
      previewHtml = previewHtml.replace(/margin:\s*[0-9.]+px\s+(auto|[0-9.]+px)(\s+[0-9.]+px)?(\s+[0-9.]+px)?/gi, 'margin: 0px');

      // 11. 一時コンテナを削除
      document.body.removeChild(tempContainer);

      console.log('[Template Preview] マスキング版プレビュー生成完了');

      return previewHtml;

    } catch (error) {
      console.error('[Template Preview] プレビュー生成エラー:', error);
      throw new Error(`プレビュー生成に失敗しました: ${error.message}`);
    }
  }

  // テンプレート保存
  async function saveCodeAsTemplate() {
    const code = app.scrapedCode;

    if (!code) {
      alert(t('templateNoCode') || '保存するコードがありません');
      return;
    }

    // デフォルトのテンプレート名を生成
    const now = new Date();
    const defaultName = `Template_${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

    // モーダルを作成
    const modal = document.createElement('div');
    modal.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 100%;
      background: rgba(0, 0, 0, 0.5);
      display: flex;
      justify-content: center;
      align-items: center;
      z-index: 10000;
    `;

    const modalContent = document.createElement('div');
    modalContent.style.cssText = `
      background: white;
      padding: 24px;
      border-radius: 8px;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
      max-width: 400px;
      width: 90%;
    `;

    modalContent.innerHTML = `
      <h3 style="margin: 0 0 16px 0; color: #333;">テンプレート名を入力</h3>
      <input type="text" id="template-name-input"
        value="${defaultName}"
        style="width: 100%; padding: 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 14px; box-sizing: border-box;"
        placeholder="テンプレート名">
      <div style="margin-top: 16px; display: flex; gap: 8px; justify-content: flex-end;">
        <button id="template-cancel-btn" style="padding: 8px 16px; background: #6c757d; color: white; border: none; border-radius: 4px; cursor: pointer;">キャンセル</button>
        <button id="template-save-btn" style="padding: 8px 16px; background: #007bff; color: white; border: none; border-radius: 4px; cursor: pointer;">保存</button>
      </div>
    `;

    modal.appendChild(modalContent);
    document.body.appendChild(modal);

    // 入力欄にフォーカス
    const input = document.getElementById('template-name-input');
    input.focus();
    input.select();

    // 保存処理
    const handleSave = async () => {
      const saveBtn = document.getElementById('template-save-btn');
      let templateName = input.value.trim();

      // 空の場合はデフォルト名を使用
      if (!templateName) {
        templateName = defaultName;
      }

      // 保存ボタンを無効化（二重クリック防止）
      if (saveBtn) {
        saveBtn.disabled = true;
        saveBtn.style.opacity = '0.6';
        saveBtn.style.cursor = 'not-allowed';
      }

      try {
        // 進捗表示を開始
        app.showProgress('テンプレートを保存中...', 120000);

        // 全自動マスキング
        const maskedCode = performMasking(code, {
          maskText: true,
          maskBullet: true,
          maskChart: true,
          randomValues: true,
          preserveSpaces: false
        });

        // マスキングしたコードからプレビューHTMLを生成
        console.log('[Template Save] マスキング版プレビューを生成中...');
        const maskedPreviewHtml = await generatePreviewFromMaskedCode(maskedCode);

        // テンプレート一覧を取得
        const templates = await getTemplates();

        // 新しいテンプレートを追加
        const newTemplate = {
          id: Date.now().toString(),
          name: templateName,
          code: maskedCode,
          previewHtml: maskedPreviewHtml, // マスキング版プレビューHTMLを保存
          originalCodeLength: code.length,
          createdAt: now.toISOString(),
          updatedAt: now.toISOString()
        };

        templates.push(newTemplate);
        await saveTemplates(templates);

        // モーダルを閉じる
        document.body.removeChild(modal);

        // 進捗表示を終了
        app.stopFakeProgress();
        app.updateProgress(100);
        setTimeout(() => {
          app.hideProgress();
          // alert(t('templateSaved') || `テンプレート「${templateName}」を保存しました！`); // アラート削除

          // テンプレート一覧モーダルを開く
          openTemplatesModal();
        }, 500);

        // 保存ボタンは非表示にしない（次回も使えるように）
        // プレビューHTMLもクリアしない（再保存できるように）
      } catch (error) {
        console.error('テンプレート保存エラー:', error);

        // 進捗表示を終了
        app.stopFakeProgress();
        app.hideProgress();

        // ボタンを再有効化
        if (saveBtn) {
          saveBtn.disabled = false;
          saveBtn.style.opacity = '1';
          saveBtn.style.cursor = 'pointer';
        }

        alert(`保存エラー: ${error.message}`);
        // モーダルは開いたまま（エラー時）
      }
    };

    // イベントリスナー
    document.getElementById('template-save-btn').addEventListener('click', handleSave);
    document.getElementById('template-cancel-btn').addEventListener('click', () => {
      document.body.removeChild(modal);
    });

    // モーダル外クリックで閉じる
    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        document.body.removeChild(modal);
      }
    });
  }

  // テンプレート一覧を取得
  async function getTemplates() {
    return new Promise((resolve) => {
      if (chrome && chrome.storage && chrome.storage.local) {
        chrome.storage.local.get(['pptxTemplates'], (result) => {
          resolve(result.pptxTemplates || []);
        });
      } else {
        resolve([]);
      }
    });
  }

  // テンプレート一覧を保存
  async function saveTemplates(templates) {
    return new Promise((resolve) => {
      if (chrome && chrome.storage && chrome.storage.local) {
        chrome.storage.local.set({ pptxTemplates: templates }, () => {
          resolve();
        });
      } else {
        resolve();
      }
    });
  }

  // テンプレートを削除
  async function deleteTemplate(id) {
    const templates = await getTemplates();
    const template = templates.find(t => t.id === id);

    if (!template) return;

    const confirmed = confirm(`テンプレート「${template.name}」を削除しますか？\nこの操作は取り消せません。`);

    if (confirmed) {
      const newTemplates = templates.filter(t => t.id !== id);
      await saveTemplates(newTemplates);
      await loadTemplatesList(); // 一覧を再読み込み
    }
  }

  // テンプレートを使用（テキスト入力UIを表示）
  async function useTemplate(id) {
    const templates = await getTemplates();
    const template = templates.find(t => t.id === id);

    if (template) {
      // テンプレート一覧を非表示にして、テキスト入力UIを表示
      showTemplateInputUI(template);
    }
  }

  // テンプレート入力UIを表示
  function showTemplateInputUI(template) {
    const templatesList = document.querySelector('#templates-list');
    if (!templatesList) return;

    // テキスト入力UIに切り替え
    templatesList.innerHTML = `
      <div style="display:flex;flex-direction:column;gap:12px;">
        <div>
          <button type="button" id="back-to-templates-btn" style="padding:8px 16px;background:#007bff;color:white;border:none;border-radius:4px;cursor:pointer;font-size:13px;font-weight:500;">
            ← 一覧に戻る
          </button>
        </div>
        <div style="display:flex;gap:16px;align-items:center;">
          ${template.previewHtml ? `
            <div style="width:240px;height:135px;flex-shrink:0;border:1px solid #ddd;border-radius:8px;overflow:hidden;background:#f9f9f9;position:relative;">
              <div style="width:240px;height:135px;position:relative;overflow:hidden;">
                <iframe
                  class="template-preview-use"
                  srcdoc="${escapeSrcdoc(template.previewHtml)}"
                  sandbox="allow-same-origin"
                  scrolling="no"
                  style="width:720px;height:405px;border:none;transform:scale(0.3333);transform-origin:top left;display:block;position:absolute;top:0;left:0;"
                ></iframe>
              </div>
            </div>
          ` : ''}
          <div style="flex:1;">
            <h3 style="margin:0 0 8px 0;font-size:16px;font-weight:600;">テンプレート: ${escapeHtml(template.name)}</h3>
            <p style="margin:0;color:#666;font-size:14px;line-height:1.5;">テンプレートを元に作るスライドの内容を入力してください</p>
          </div>
        </div>
        <textarea id="template-input-text" style="width:100%;height:200px;padding:12px;border:1px solid #ddd;border-radius:4px;font-size:14px;resize:vertical;box-sizing:border-box;" placeholder="ここに内容を入力してください..."></textarea>
        <div style="display:flex;justify-content:flex-end;gap:8px;">
          <button type="button" id="template-create-btn" style="padding:10px 24px;background:#bf0000;color:white;border:none;border-radius:4px;cursor:pointer;font-size:14px;font-weight:500;">
            作成
          </button>
        </div>
      </div>
    `;

    // 戻るボタンのイベントリスナー
    const backBtn = document.querySelector('#back-to-templates-btn');
    if (backBtn) {
      backBtn.onclick = () => {
        loadTemplatesList();
      };
    }

    // 作成ボタンのイベントリスナー
    const createBtn = document.querySelector('#template-create-btn');
    if (createBtn) {
      createBtn.onclick = async () => {
        const inputText = document.querySelector('#template-input-text').value.trim();
        if (!inputText) {
          alert('テキストを入力してください');
          return;
        }
        await createWithTemplate(template, inputText);
      };
    }
  }

  // テンプレートとテキストを組み合わせてAIに送信
  async function createWithTemplate(template, inputText) {
    try {
      // モーダルを閉じる
      closeTemplatesModal();

      // テンプレートコードとユーザー入力を組み合わせたプロンプトを作成
      const combinedPrompt = `${inputText}\n\n以下のPptxgenjsコードをベースに使用してください：\n\`\`\`javascript\n${template.code}\n\`\`\``;

      // AIチャットに送信
      await handOff({ prompt: combinedPrompt }, app.CHAT_URL);
    } catch (error) {
      console.error('[Template Use] エラー:', error);
      alert(`エラーが発生しました: ${error.message}`);
    }
  }

  // 日付フォーマット
  function formatTemplateDate(dateString) {
    const date = new Date(dateString);
    const now = new Date();
    const diff = now - date;
    const days = Math.floor(diff / (1000 * 60 * 60 * 24));

    if (days === 0) {
      return '今日';
    } else if (days === 1) {
      return '昨日';
    } else if (days < 7) {
      return `${days}日前`;
    } else {
      return date.toLocaleDateString('ja-JP', { year: 'numeric', month: '2-digit', day: '2-digit' });
    }
  }

  // HTMLエスケープ（テンプレート名用）
  function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  }

  // srcdoc用HTMLエスケープ
  function escapeSrcdoc(html) {
    if (!html) return '';
    return html
      .replace(/&/g, '&amp;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  // テンプレート一覧を読み込む
  async function loadTemplatesList() {
    const templates = await getTemplates();
    const templatesList = document.querySelector('#templates-list');

    if (!templatesList) return;

    if (templates.length === 0) {
      templatesList.innerHTML = `
        <div class="no-templates">
          <p style="font-size:48px;margin-bottom:16px;">📝</p>
          <p style="font-size:16px;">保存されたテンプレートはありません</p>
        </div>
      `;
      return;
    }

    // 新しい順にソート
    const sortedTemplates = templates.sort((a, b) =>
      new Date(b.createdAt) - new Date(a.createdAt)
    );

    templatesList.innerHTML = sortedTemplates.map(template => `
      <div class="template-item" data-id="${template.id}">
        <div class="template-thumbnail-container">
          ${template.previewHtml ? `
            <iframe
              class="template-thumbnail"
              srcdoc="${escapeSrcdoc(template.previewHtml)}"
              sandbox="allow-same-origin"
              scrolling="no"
            ></iframe>
          ` : `
            <div class="template-thumbnail-placeholder">
              <span style="font-size:64px;">📊</span>
            </div>
          `}
        </div>
        <div class="template-content">
          <div class="template-header">
            <h3 class="template-name">${escapeHtml(template.name)}</h3>
            <span class="template-date">${formatTemplateDate(template.createdAt)}</span>
          </div>
          <div class="template-info">
            <span>📝 ${template.code.length} 文字</span>
          </div>
          <div class="template-actions">
            <button class="use-template-btn" data-id="${template.id}" data-i18n="templateUse">使用する</button>
            <button class="delete-template-btn" data-id="${template.id}" data-i18n="templateDelete">削除</button>
          </div>
        </div>
      </div>
    `).join('');

    // 翻訳を適用
    applyTranslations(templatesList);

    // イベントリスナーを追加
    document.querySelectorAll('.use-template-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = e.target.dataset.id;
        useTemplate(id);
      });
    });

    document.querySelectorAll('.delete-template-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = e.target.dataset.id;
        deleteTemplate(id);
      });
    });
  }

  // テンプレートモーダルを閉じる
  function closeTemplatesModal() {
    const templatesModal = document.querySelector('#templates-modal');
    if (templatesModal) {
      console.log('[closeTemplatesModal] Closing modal');
      templatesModal.setAttribute('aria-hidden', 'true');
      templatesModal.style.display = 'none';
    }
  }

  // テンプレートモーダルを開く（イベントリスナーも設定）
  function openTemplatesModal() {
    const templatesModal = document.querySelector('#templates-modal');
    if (templatesModal) {
      console.log('[openTemplatesModal] Opening modal');
      templatesModal.setAttribute('aria-hidden', 'false');
      templatesModal.style.display = 'flex';

      // 閉じるボタンのイベントリスナーを設定（毎回設定し直す）
      const closeBtn = templatesModal.querySelector('.templates-modal-close');
      if (closeBtn) {
        console.log('[openTemplatesModal] Setting up close button');

        // 既存のイベントリスナーをクリア（onclickを上書き）
        closeBtn.onclick = null;

        // 新しいイベントリスナーを設定
        closeBtn.onclick = (e) => {
          console.log('[Templates Modal] Close button clicked via onclick');
          e.preventDefault();
          e.stopPropagation();
          closeTemplatesModal();
        };

        // ホバー効果を確認するためのログ
        closeBtn.onmouseenter = () => {
          console.log('[Templates Modal] Close button hovered');
        };
      } else {
        console.warn('[openTemplatesModal] Close button not found!');
      }

      // モーダル背景クリックで閉じる
      templatesModal.onclick = (e) => {
        if (e.target === templatesModal) {
          console.log('[Templates Modal] Background clicked');
          closeTemplatesModal();
        }
      };

      loadTemplatesList();
    } else {
      console.warn('[openTemplatesModal] Modal not found!');
    }
  }

  // ヘッダー内のボタン表示・状態を更新
  function updateHeaderButtons() {
    try {
      const showMulti = app.multiSlideMode;
      const htmlFound  = extractAllHtmlBlocks().length > 0 || !!app.lastHtml;
      const pptxFound  = extractAllPptxSnippets().length > 0 || !!app.lastPptx;
      const jsonFound  = hasSlidesJson();
      const idsSingle = ['convert-btn', 'download-btn', 'download-api-btn', 'preview-pptx-btn'];
      const idsMulti  = ['create-pptx-btn', 'download-multi-btn', 'start-all-btn'];
      idsSingle.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        if (showMulti) {
          el.style.display = 'none';
        } else if (id === 'download-btn' || id === 'download-api-btn') {
          el.style.display = pptxFound ? 'inline-flex' : 'none';
          if (id === 'download-api-btn') {
            el.disabled = !pptxFound;
            el.classList.toggle('active', pptxFound);
          }
        } else if (id === 'preview-pptx-btn') {
          el.style.display = pptxFound ? 'inline-flex' : 'none';
          el.disabled = !pptxFound;
          el.classList.toggle('active', pptxFound);

          // Update button text based on whether preview is already showing
          const textSpan = el.querySelector('span[data-i18n]');
          if (textSpan) {
            if (app.currentPreview === 'pptx') {
              // Preview is active - show "Update Preview"
              textSpan.setAttribute('data-i18n', 'previewPptxUpdate');
              textSpan.textContent = t('previewPptxUpdate');
            } else {
              // No preview or different preview - show "Preview"
              textSpan.setAttribute('data-i18n', 'previewPptx');
              textSpan.textContent = t('previewPptx');
            }
          }
        } else if (id === 'convert-btn') {
          el.style.display = htmlFound ? 'inline-flex' : 'none';
        } else {
          el.style.display = 'inline-flex';
        }
      });
      idsMulti.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        if (id === 'start-all-btn') {
          const show = showMulti && jsonFound && app.currentPreview === 'json';
          el.style.display = show ? 'inline-flex' : 'none';
          el.classList.toggle('flashy', show);
        } else if (id === 'create-pptx-btn') {
          const show = showMulti && htmlFound;
          el.style.display = show ? 'inline-flex' : 'none';
        } else if (id === 'download-multi-btn') {
          const show = showMulti && pptxFound;
          el.style.display = show ? 'inline-flex' : 'none';
        } else {
          el.style.display = showMulti ? 'inline-flex' : 'none';
        }
      });

      // Handle save-preview-btn - always visible, enable when Pptxgenjs code detected
      const savePreviewBtn = document.getElementById('save-preview-btn');
      if (savePreviewBtn) {
        savePreviewBtn.disabled = !pptxFound;
        savePreviewBtn.classList.toggle('active', pptxFound);
      }
    } catch (e) {
      console.error('updateHeaderButtons error', e);
    }
  }

  // 各ボタンの活性状態を判定して切り替える
  function updateActionStates(htmlFound, pptxFound, jsonFound) {
    const createBtn = document.getElementById('create-pptx-btn');
    if (createBtn) {
      createBtn.disabled = !htmlFound;
      createBtn.classList.toggle('active', htmlFound);
    }
    const startBtn = document.getElementById('start-all-btn');
    if (startBtn) {
      const creating = app.creatingSlides;
      startBtn.disabled = !jsonFound || creating;
      startBtn.classList.toggle('active', jsonFound && !creating);
      startBtn.classList.toggle('flashy', jsonFound && !creating);
      if (creating) {
        startBtn.dataset.i18n = 'creatingSlides';
        startBtn.textContent = t('creatingSlides');
      } else {
        startBtn.dataset.i18n = 'startSlides';
        startBtn.textContent = startBtn.dataset.orig || t('startSlides');
        if (jsonFound && app.currentPreview === 'json') {
          showStartArrow(startBtn);
        }
      }
    }
    const dlMulti = document.getElementById('download-multi-btn');
    if (dlMulti) {
      dlMulti.disabled = !pptxFound;
      dlMulti.classList.toggle('active', pptxFound);
    }
  }

  // スタートボタンの位置を矢印で示す
  function showStartArrow(btn) {
    try {
      if (app.arrowTimer !== null) return;
      const panel = document.getElementById(app.PANEL_ID);
      const rect = panel?.getBoundingClientRect();
      const arrow = document.createElement('img');
      arrow.id = 'start-arrow';
      arrow.src = safeGetURL(`icon/solid/arrow-up.svg?color=${PRIMARY_COLOR.slice(1)}`);
      if (rect) {
        Object.assign(arrow.style, {
          left: `${rect.left + 70}px`,
          top: `${rect.top + 70}px`,
        });
      }
      document.body.appendChild(arrow);
      const overlay = panel?.querySelector('#arrow-overlay');
      if (overlay) overlay.style.display = 'block';
      app.arrowTimer = setTimeout(() => {
        arrow.remove();
        if (overlay) overlay.style.display = 'none';
        app.arrowTimer = -1;
      }, 2000);
    } catch (e) {
      console.error('showStartArrow error', e);
    }
  }


  // 登録ボタンを完了状態の見た目にする
  function markRegButtonDone(btn) {
    if (!btn) return;
    btn.textContent = '登録完了！';
    btn.style.background = 'none';
    btn.style.boxShadow = 'none';
    btn.style.color = PRIMARY_COLOR;
  }

  // 登録後のUI要素を遅れて非表示にする
  function scheduleHideRegElements(justRegistered = false) {
    if (app.regHideTimer) return;
    if (!(app.previewRegDone && app.pptxRegDone)) return;
    if (app.templateModalOpen) return;
    app.regHideTimer = setTimeout(() => {
      if (app.templateModalOpen) {
        app.regHideTimer = null;
        return;
      }
      app.regHideTimer = null;
      if (justRegistered) {
        location.reload();
        return;
      }
      if (app.isTyping) {
        app.pendingRender = true;
      } else {
        renderPreview();
      }
    }, 2000);
  }

  // ダウンロード処理を中止する
  function cancelDownload() {
    if (injector.timer) {
      clearInterval(injector.timer);
      injector.timer = null;
    }
    if (app.autoDlTimer) {
      clearTimeout(app.autoDlTimer);
      app.autoDlTimer = null;
    }
    if (injector.isProcessing) {
      injector.shouldStop = true;
    }
    if (typeof app.cancelDownload === 'function') {
      app.cancelDownload();
    }
    hideProgress();
  }

  // 生成済みHTMLをチャットに送るモードに切り替える
  function toggleChat() {
    const btn = document.getElementById('convert-btn');
    if (btn) {
      btn.disabled = true;
      btn.classList.remove('active');
    }
    showProgress('creatingPptx', 120000);
    const data = {
      [app.PREVIEW_PARAM]: '1',
      [app.AUTO_DL_PARAM]: '1',
    };
    if (app.lastHtml) {
      data[app.HTML_PARAM] = LZString.compressToEncodedURIComponent(app.lastHtml);
    }
    handOff(data, app.CHAT_URL);
  }

  // HTMLからPPTXを生成する処理を開始
  function startHtmlSlides() {
    const btn = document.getElementById('create-pptx-btn');
    if (btn) {
      btn.disabled = true;
      btn.classList.remove('active');
    }
    app.logStart = Date.now();
    logStep('HTML スライド送信を開始');
    const frames = extractPreviewFrames();
    // collect full HTML for each slide using srcdoc (fallback to DOM)
    let slides = frames
      .map(f => f?.srcdoc || f?.contentDocument?.documentElement?.outerHTML || '')
      .filter(Boolean);
    if (!slides.length && Array.isArray(app.htmlSlides) && app.htmlSlides.length) {
      slides = app.htmlSlides;
    }
    if (!slides.length) {
      hideProgress();
      return;
    }
    app.htmlSlides = slides;
    showProgress('creatingPptx', 120000);
    const data = {
      [app.PREVIEW_PARAM]: '1',
      [app.AUTO_DL_PARAM]: '1',
      [app.HTML_LIST_PARAM]: slides.map(h => LZString.compressToEncodedURIComponent(h)),
      logStart: app.logStart,
    };
    handOff(data, app.CHAT_URL);
  }

  // ホーム画面の表示切替
  function toggleHome() {
    app.showHome = !app.showHome;
    const btn = document.getElementById('home-btn');
    if (btn) {
      btn.classList.toggle('active', app.showHome);
      const img = btn.querySelector('img');
      if (img) img.src = safeGetURL(app.showHome ? app.HOME_WHITE_SRC : app.HOME_GRAY_SRC);
    }
    renderPreview();
  }

  // startChat moved to aiClient.js

  // startPromptList moved to aiClient.js

  // autoSendPreviewHtml moved to aiClient.js

  // テキストをクリップボードへコピー
  async function copyToClipboard(text) {
    try {
      await navigator.clipboard.writeText(text);
      return true;
    } catch (e) {
      try {
        const ta = document.createElement('textarea');
        ta.value = text;
        ta.style.position = 'fixed';
        ta.style.opacity = '0';
        document.body.appendChild(ta);
        ta.focus();
        ta.select();
        const ok = document.execCommand('copy');
        ta.remove();
        return ok;
      } catch {
        return false;
      }
    }
  }

  // 最後に生成したHTMLをコピー
  function copyHtml() {
    if (!app.lastHtml) return;
    navigator.clipboard.writeText(app.lastHtml).then(() => {
      const btn = document.getElementById('copy-html-btn');
      if (btn) {
        btn.dataset.i18n = 'copied';
        btn.textContent = t('copied');
        btn.disabled = true;
        btn.classList.remove('active');
      }
    }).catch(() => {
      alert(t('copyFailedAlert'));
    });
  }

  // 現在扱っているHTML文字列を取得
  function getCurrentHtmlText() {
    try {
      if (app.multiSlideMode) {
        const all = extractAllHtmlBlocks();
        if (all && all.length) return all.join('\n');
        if (app.htmlSlides && app.htmlSlides.length) return app.htmlSlides.join('\n');
      }
      const { html } = extractBlocks();
      return html || app.lastHtml || '';
    } catch (e) {
      console.error('getCurrentHtmlText error', e);
      return app.lastHtml || '';
    }
  }


  // 一定時間後に自動ダウンロードを実行
  function scheduleAutoDownload(delay = 120000, autoDownload = true) {
    if (app.autoDlTimer) clearTimeout(app.autoDlTimer);
    app.autoDlTimer = setTimeout(() => {
      app.autoDlTimer = null;
      try {
        if (app.isTyping) {
          app.pendingRender = true;
          return;
        }
        updateProgressMessage('preparingExport');
        const { pptx } = extractBlocks();
        if (!pptx) return;
        app.scrapedCode = pptx;
        if (document.getElementById(app.PANEL_ID)) {
          renderPreview();
        }
        if (autoDownload) {
          let allSnippets = [];
          if (typeof extractAllPptxSnippets === 'function') {
            allSnippets = extractAllPptxSnippets();
          }
          const multi =
            app.multiSlideMode || (Array.isArray(allSnippets) && allSnippets.length > 1);
          if (multi && typeof downloadMultiplePptx === 'function') {
            downloadMultiplePptx();
          } else if (app.downloadPptx) {
            app.downloadPptx();
          }
        }
      } catch (e) {
        console.error('auto download failed', e);
      }
    }, delay);
  }

  // ページからPPTXコードを検出してプレビュー
  function detectPptxAndRender() {
    try {
      const { pptx } = extractBlocks();
      if (!pptx) return;
      app.scrapedCode = pptx;
      if (document.getElementById(app.PANEL_ID)) {
        renderPreview();
      }
    } catch (e) {
      console.error('pptx detection failed', e);
    }
  }

  // ページ内容が出揃うまで定期的にチェック
  function monitorHtmlRender(timeout = 600000, interval = 1000) {
    if (app.htmlMonitorTimer) return;
    const start = Date.now();
    app.htmlMonitorTimer = setInterval(() => {
      try {
        const { html, pptx } = extractBlocks();
        const jsonFound = hasSlidesJson();
        if (jsonFound && !app.userSlideModeSet) app.multiSlideMode = true;
        if (html || pptx || jsonFound) {
          clearInterval(app.htmlMonitorTimer);
          app.htmlMonitorTimer = null;
          if (document.getElementById(app.PANEL_ID)) {
            renderPreview();
          }
        } else if (Date.now() - start > timeout) {
          clearInterval(app.htmlMonitorTimer);
          app.htmlMonitorTimer = null;
        }
      } catch (e) {
        console.error('monitorHtmlRender error', e);
        clearInterval(app.htmlMonitorTimer);
        app.htmlMonitorTimer = null;
      }
    }, interval);
  }

  // Extract the latest JSON slides definition from the page
  // ページ内のJSONコードからスライド情報を取り出す
  async function extractSlidesJson() {
    try {
      const nodes = document.querySelectorAll('code.language-json');
      if (!nodes.length) throw new Error('JSON code block not found');
      const latest = nodes[nodes.length - 1];
      const text = (latest.textContent || latest.innerText || '').trim();
      const data = JSON.parse(text);
      if (!Array.isArray(data.slides)) {
        throw new Error('slides array missing');
      }
      return data.slides;
    } catch (e) {
      console.error('extractSlidesJson error', e);
      return null;
    }
  }

  // JSON形式のスライドが存在するか判定
  function hasSlidesJson() {
    try {
      const nodes = document.querySelectorAll('code.language-json');
      for (let i = nodes.length - 1; i >= 0; i--) {
        const txt = (nodes[i].textContent || nodes[i].innerText || '').trim();
        if (!txt) continue;
        try {
          const data = JSON.parse(txt);
          if (Array.isArray(data.slides)) return true;
        } catch {
          continue;
        }
      }
      return false;
    } catch (e) {
      console.error('hasSlidesJson error', e);
      return false;
    }
  }

  // 最新のJSONスライドデータを取得
  function getSlidesJson() {
    try {
      const nodes = document.querySelectorAll('code.language-json');
      if (!nodes.length) return null;
      const latest = nodes[nodes.length - 1];
      const text = (latest.textContent || latest.innerText || '').trim();
      const data = JSON.parse(text);
      if (!Array.isArray(data.slides)) return null;
      return data.slides;
    } catch (e) {
      console.error('getSlidesJson error', e);
      return null;
    }
  }
  // プレビュー用iframeをすべて取得
  function extractPreviewFrames() {
    try {
      return Array.from(document.querySelectorAll('.slide-frame iframe, .preview-slide iframe'));
    } catch (e) {
      console.error('extractPreviewFrames error', e);
      return [];
    }
  }

  // プレビューのスライドをウィンドウに合わせて縮小
  function fitPreviewSlides() {
    try {
      const cont = document.querySelector(`#${app.PANEL_ID} .preview-container`);
      if (!cont) return;
      const w = cont.clientWidth;
      if (!w) return;
      const h = w * 720 / 1280;
      cont.querySelectorAll('.preview-slide').forEach(slide => {
        slide.style.width = `${w}px`;
        slide.style.height = `${h}px`;
        const inner = slide.querySelector('iframe, .json-slide');
        if (inner) inner.style.transform = `scale(${w / 1280})`;
      });
      fitJsonPromptText();
    } catch (e) {
      console.error('fitPreviewSlides error', e);
    }
  }

  // JSONスライドのプロンプト文字サイズを調整
  function fitJsonPromptText() {
    try {
      const slides = document.querySelectorAll(`#${app.PANEL_ID} .json-slide`);
      slides.forEach(slide => {
        const title = slide.querySelector('.json-slide-title');
        const prompt = slide.querySelector('.json-slide-prompt');
        if (!prompt) return;

        const titleStyles = title ? getComputedStyle(title) : null;
        const titleMargin = titleStyles ? parseFloat(titleStyles.marginBottom) || 0 : 0;
        const titleHeight = title ? title.offsetHeight + titleMargin : 0;

        const maxHeight = 720 - 80 - titleHeight;
        let fontSize = parseFloat(getComputedStyle(prompt).fontSize) || 24;

        while (fontSize > 10 && prompt.scrollHeight > maxHeight) {
          fontSize -= 1;
          prompt.style.fontSize = `${fontSize}px`;
        }
      });
    } catch (e) {
      console.error('fitJsonPromptText error', e);
    }
  }

  // チャット入力欄にテキストを挿入
  function insertPromptText(text) {
    try {
      const box = document.querySelector('div[contenteditable="true"][role="textbox"]') ||
                  document.querySelector('#start-text');
      if (!box) return;
      if (box.tagName === 'TEXTAREA') {
        box.value = text;
      } else {
        box.textContent = text;
      }
      box.dispatchEvent(new Event('input', { bubbles: true }));
      box.focus();
    } catch (e) {
      console.error('insertPromptText error', e);
    }
  }

  // エラー内容を含むテキストを入力欄に挿入
  function insertErrorPrompt(errMsg) {
    const text =
      `エラーが起きてパワーポイントが適切にエクスポートできないので、コードを適切に書き換えてエラーを解決して\n${errMsg || ''}`;
    let attempts = 0;
    const maxAttempts = 5;
    const tryInsert = () => {
      try {
        insertPromptText(text);
        const box = document.querySelector('div[contenteditable="true"][role="textbox"], #start-text, textarea');
        const content = box ? (box.tagName === 'TEXTAREA' ? box.value : box.textContent) : '';
        if (!box || content !== text) throw new Error('not inserted');
      } catch {
        if (attempts++ < maxAttempts) {
          setTimeout(tryInsert, 500);
        }
      }
    };
    tryInsert();
  }

  // 提案プロンプトのリストを更新
  function updatePromptCandidates() {
    try {
      const wrap = document.getElementById(app.PANEL_ID);
      if (!wrap) return;
      const existing = wrap.querySelector('#prompt-candidates');
      const previewType = app.currentPreview;
      const frames = extractPreviewFrames();
      let pptxDetected = false;
      if (previewType === 'html' || previewType === 'pptx') {
        pptxDetected = frames.some(f => {
          try { return !!f.contentWindow?.PptxGenJS; } catch { return false; }
        });
      }
      let promptKeys = null;
      let promptType = null;
      if (previewType === 'html' && !pptxDetected) {
        promptType = 'html';
        promptKeys = [
          'promptSuggestion1',
          'promptSuggestion2',
          'promptSuggestion3',
        ];
      } else if (pptxDetected || previewType === 'pptx') {
        promptType = 'pptx';
        promptKeys = [
          'pptxPromptSuggestion1',
          'pptxPromptSuggestion2',
          'pptxPromptSuggestion3',
          'pptxPromptSuggestion4',
        ];
      }
      if (promptKeys) {
        const footer = wrap.querySelector('.preview-footer');
        const bottom = footer ? footer.offsetHeight + 8 : 56;
        let div = existing;
        if (!div || div.dataset.type !== promptType) {
          if (div) div.remove();
          div = document.createElement('div');
          div.id = 'prompt-candidates';
          div.dataset.type = promptType;
          const header = document.createElement('div');
          header.className = 'prompt-header';
          const label = document.createElement('div');
          label.className = 'prompt-label';
          label.setAttribute('data-i18n', 'promptSuggestionsLabel');
          label.textContent = t('promptSuggestionsLabel');
          header.appendChild(label);
          const toggle = document.createElement('div');
          toggle.className = 'prompt-toggle';
          toggle.textContent = app.promptCollapsed ? '^' : '×';
          const togglePanel = () => {
            app.promptCollapsed = !app.promptCollapsed;
            div.classList.toggle('collapsed', app.promptCollapsed);
            toggle.textContent = app.promptCollapsed ? '^' : '×';
          };
          toggle.onclick = e => {
            e.stopPropagation();
            togglePanel();
          };
          header.appendChild(toggle);
          div.appendChild(header);
          div.addEventListener('click', e => {
            if (!e.target.classList.contains('prompt-option') && !e.target.classList.contains('prompt-toggle')) {
              togglePanel();
            }
          });
          const list = document.createElement('div');
          list.className = 'prompt-options';
          promptKeys.forEach((key, i) => {
            const box = document.createElement('div');
            box.className = `prompt-option option-${i + 1}`;
            box.setAttribute('data-i18n', key);
            box.textContent = t(key);
            box.onclick = e => {
              e.stopPropagation();
              insertPromptText(t(key));
            };
            list.appendChild(box);
          });
          div.appendChild(list);
          wrap.appendChild(div);
        }
        div.style.bottom = `${bottom}px`;
        div.classList.toggle('collapsed', app.promptCollapsed);
      } else if (existing) {
        existing.remove();
      }
    } catch (e) {
      console.error('updatePromptCandidates error', e);
    }
    adjustPreviewHeight();
  }

  // プレビュー領域の高さをヘッダー・フッターに合わせ調整
  function adjustPreviewHeight() {
    try {
      const wrap = document.getElementById(app.PANEL_ID);
      if (!wrap) return;
      const header = wrap.querySelector('.preview-header');
      const footer = wrap.querySelector('.preview-footer');
      const cont   = wrap.querySelector('.preview-container');
      if (!header || !footer || !cont) return;
      const total = header.offsetHeight + footer.offsetHeight;
      cont.style.maxHeight = `calc(100vh - ${total}px)`;
    } catch (e) {
      console.error('adjustPreviewHeight error', e);
    }
  }

  // プレビュー中のHTMLをまとめて送信
  async function sendHtmlSlides() {
    try {
      let slides = [];
      let frames = extractPreviewFrames();
      const start = Date.now();
      // wait up to 10s for frames to appear
      while (!frames.length && Date.now() - start < 10000) {
        await sleep(200);
        frames = extractPreviewFrames();
      }
      if (frames.length) {
        slides = frames
          .map(f => f?.srcdoc || f?.contentDocument?.documentElement?.outerHTML || '')
          .filter(Boolean);
      } else if (Array.isArray(app.htmlSlides) && app.htmlSlides.length) {
        slides = app.htmlSlides;
      }

      if (!slides.length) return;

      injector.config = {
        messages: [],
        selectors: {
          textbox: 'div[contenteditable="true"][role="textbox"]',
          button: 'button[type="submit"], button:last-of-type'
        },
        delays: {
          beforeInput: 100,
          afterInput: 100,
          beforeSend: 100,
          afterSend: 50,
          betweenMessages: 100
        },
        useEnterKey: true,
        debug: false
      };

      injector.timer = setInterval(async () => {
        if (injector.isProcessing) return;
        const box = findElement(injector.config.selectors.textbox);
        if (!box) return;
        clearInterval(injector.timer);
        injector.timer = null;
        injector.isProcessing = true;
        showProgress('creatingPptx', 120000);

        try {
          await ensureResearchOff();
          await ensureWebSearchOff();
          for (let i = 0; i < slides.length; i++) {
            logStep(`HTMLスライド${i + 1}を送信`);
            const html = slides[i];
            const msg = CONVERT_PREFIX + '\n\n' + html;
            await sendSingleMessage(msg, false, false);
            await waitForAnswerComplete();
            logStep('回答を受信しました');
            await sleep(injector.config.delays.betweenMessages);
          }
          logStep('HTML スライド送信が完了しました');
        } catch (e) {
          console.error('sendHtmlSlides error', e);
        } finally {
          injector.isProcessing = false;
          hideProgress();
          try {
            await waitForPptxSnippets(slides.length);
            downloadMultiplePptx();
          } catch (e) {
            console.error('auto multi download failed', e);
          }
        }
      }, 100);
    } catch (e) {
      console.error('sendHtmlSlides error', e);
    }
  }

  // JSON情報から一括でスライド生成を開始
  async function startBatchCreation() {
    if (app.creatingSlides) return;
    app.logStart = Date.now();
    logStep('開始ボタンが押されました');
    app.creatingSlides = true;
    const btn = document.getElementById('start-all-btn');
    if (btn) {
      btn.dataset.orig = btn.textContent;
      btn.disabled = true;
      btn.classList.remove('active');
      btn.dataset.i18n = 'creatingSlides';
      btn.textContent = t('creatingSlides');
    }
    try {
      await sendSlidesFromJson();
      app.htmlSlides = extractAllHtmlBlocks();
      startHtmlSlides();
    } catch (e) {
      console.error('startBatchCreation error', e);
      if (btn) {
        btn.disabled = false;
        btn.classList.add('active');
        btn.dataset.i18n = 'startSlides';
        btn.textContent = btn.dataset.orig || t('startSlides');
      }
    } finally {
      app.creatingSlides = false;
    }
  }

  // ページの変化を監視しプレビューを更新
  function startObserveChanges() {
    const shouldIgnore = node => {
      if (!node) return false;
      const el = node.nodeType === Node.TEXT_NODE ? node.parentElement : node;
      return el && el.closest(`#${app.PANEL_ID}, #${app.TOGGLE_ID}`);
    };

    app.observer = new MutationObserver(mutations => {
      if (mutations.some(m => shouldIgnore(m.target))) return;
      if (app.downloading) return;
      clearTimeout(app.renderTimer);
      if (app.isTyping) {
        app.pendingRender = true;
        return;
      }
      app.renderTimer = setTimeout(renderPreview, 300);
    });
    app.observer.observe(document.body, {
      childList: true,
      subtree: true,
      characterData: true,
    });
  }

  initAIClient(app, {
    handOff,
    showWaitOverlay,
    showProgress,
    updateProgressMessage,
    scheduleAutoDownload,
    monitorHtmlRender,
    extractSlidesJson,
    resetSearchMode,
    t,
    CONVERT_PREFIX,
    JSON_TO_HTML_PREFIX,
  });

  Object.assign(app, {
    injectStyles,
    injectToggle,
    togglePanel,
    closePanel,
    openPanel,
    extractBlocks,
    extractAllHtmlBlocks,
    extractAllPptxSnippets,
    generateMultipleSlides,
    downloadMultiplePptx,
    renderPreview,
    showProgress,
    hideProgress,
    updateProgressMessage,
    updateProgress,
    startFakeProgress,
    stopFakeProgress,
    showWaitOverlay,
    hideWaitOverlay,
    toggleHome,
    toggleChat,
    startChat,
    autoSendPreviewHtml,
    autoSendPrompt,
    scheduleAutoDownload,
    detectPptxAndRender,
    extractSlidesJson,
    hasSlidesJson,
    sendSlidesFromJson,
    extractPreviewFrames,
    sendHtmlSlides,
    startHtmlSlides,
    cancelDownload,
    copyHtml,
    startObserveChanges,
    ensureResearchOn,
    ensureResearchOff,
    monitorDeepResearchLoading,
    monitorHtmlRender,
    resetSearchMode,
    startBatchCreation,
  });

  window.previewApp = app;
  window.previewApp.insertPromptText = insertPromptText;
  window.previewApp.insertErrorPrompt = insertErrorPrompt;
  window.previewApp.sendHtmlSlides = sendHtmlSlides;
  window.previewApp.extractPreviewFrames = extractPreviewFrames;
  window.previewApp.downloadMultiplePptx = downloadMultiplePptx;
  window.previewApp.startBatchCreation = startBatchCreation;
  document.dispatchEvent(new Event('previewAppReady'));

  // DOM準備完了後の初期処理
  function onReady() {
    const p = app.payload || {};
    loadRegStatus(app, () => {
      loadPanelState(app, () => {
        loadAnnouncementState(app, (state) => {
          syncAnnouncementState(state);
          const urlMode = new URLSearchParams(location.search).get(app.SEARCH_PARAM);
          if (urlMode === 'websearch' || urlMode === 'web') {
            app.searchMode = 'web';
          } else if (urlMode === 'deepresearch' || urlMode === 'deep') {
            app.searchMode = 'deep';
          } else {
            app.searchMode =
              p[app.SEARCH_PARAM] || (p[app.DEEP_PARAM] === '1' ? 'deep' : app.searchMode);
          }
          app.slidePrompt = p[app.SLIDE_PARAM] || '';
          app.injectToggle();
          const urlWait = new URLSearchParams(location.search).get(app.WAIT_PARAM);
          if (p[app.WAIT_PARAM] === '1' || urlWait === '1') {
            showWaitOverlay();
          }
          autoSendPreviewHtml();
          autoSendPrompt();
          monitorHtmlRender();
          if (app.htmlSlides && app.htmlSlides.length) {
            sendHtmlSlides().catch(e => console.error(e));
          }
          adjustPreviewHeight();
          window.addEventListener('resize', adjustPreviewHeight);
        });
      });
    });
  }

  document.readyState === 'loading'
    ? window.addEventListener('DOMContentLoaded', onReady)
    : onReady();
})();
