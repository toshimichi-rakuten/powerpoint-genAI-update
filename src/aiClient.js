export const injector = {
  isProcessing: false,
  shouldStop: false,
  config: null,
  timer: null,
};

let app;
let helpers = {};

// アプリ本体と補助関数を受け取り初期設定する
export function initAIClient(appRef, injectedHelpers = {}) {
  app = appRef;
  helpers = injectedHelpers;
}

// デバッグ時のみログを出力する内部関数
function injLog(...args) {
  if (injector.config?.debug) {
    console.log('[TextInjector]', ...args);
  }
}

// 処理の経過時間と共にメッセージを表示する
export function logStep(message) {
  if (!app.logStart) app.logStart = Date.now();
  const sec = ((Date.now() - app.logStart) / 1000).toFixed(1);
  console.log(`[${sec}s] ${message}`);
}


// ページ内で表示中の要素をセレクタで探す
export function findElement(selector) {
  try {
    const list = document.querySelectorAll(selector);
    for (const el of list) {
      if (el.offsetParent !== null) return el;
    }
  } catch (e) {
    injLog('selector error', e);
  }
  return null;
}

// 指定ミリ秒だけ待機する
export function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

function isVisible(el) {
  if (!el) return false;
  if (el.offsetParent !== null) return true;
  const rect = el.getBoundingClientRect?.();
  return !!rect && rect.width > 0 && rect.height > 0;
}

function findComposerRoot() {
  const composer = document.querySelector('[contenteditable="true"][role="textbox"]');
  if (!composer) return null;
  return composer.closest('form') || composer.parentElement;
}

function findWebSearchChip(root) {
  if (!root) return null;
  const candidates = Array.from(root.querySelectorAll('span, div, button')).filter(isVisible);
  for (const el of candidates) {
    const text = (el.textContent || '').toLowerCase();
    if (!text.includes('web search') && !text.includes('ウェブ検索')) continue;
    const chip = el.closest('div');
    if (!chip) continue;
    if (!isVisible(chip)) continue;
    if (!root.contains(chip)) continue;
    const closeBtn = chip.querySelector('button[aria-label], button');
    if (!isVisible(closeBtn)) continue;
    const aria = (closeBtn.getAttribute('aria-label') || '').toLowerCase();
    const textLabel = (closeBtn.textContent || '').toLowerCase();
    if (aria) {
      if (!aria.includes('close') && !aria.includes('閉')) continue;
    } else if (!textLabel.includes('close') && !textLabel.includes('閉')) {
      continue;
    }
    return { chip, closeBtn };
  }
  return null;
}

async function tryCloseWebSearchChip({ searchDuration = 500, confirmDuration = 500 } = {}) {
  const composerRoot = findComposerRoot();
  if (!composerRoot) return false;

  const deadline = Date.now() + searchDuration;
  while (Date.now() < deadline) {
    const chip = findWebSearchChip(composerRoot);
    if (chip) {
      console.log('[ensureWebSearchOff] closing textbox chip');
      chip.closeBtn.click();
      const confirmDeadline = Date.now() + confirmDuration;
      while (Date.now() < confirmDeadline) {
        await sleep(100);
        if (!findWebSearchChip(composerRoot)) {
          console.log('[ensureWebSearchOff] confirmed textbox chip removed');
          return true;
        }
      }
      console.log('[ensureWebSearchOff] chip still visible after close attempt');
      return false;
    }
    await sleep(100);
  }
  return false;
}

// Web検索ボタンが有効か調べる
function isWebSearchActive(btn) {
  return (
    btn.getAttribute('aria-pressed') === 'true' ||
    btn.classList.contains('bg-tinted-blue') ||
    btn.dataset.colorSchema === 'blue' ||
    btn.classList.contains('active')
  );
}

// Researchボタンが有効か調べる
function isResearchActive(btn) {
  return (
    btn.getAttribute('aria-pressed') === 'true' ||
    btn.classList.contains('bg-tinted-purple') ||
    btn.dataset.colorSchema === 'purple' ||
    btn.classList.contains('active')
  );
}

// Web検索ボタンを確実にオフにする
export async function ensureWebSearchOff(maxAttempts = 5, interval = 100) {
  const chipClosed = await tryCloseWebSearchChip();
  if (chipClosed) return;

  for (let i = 0; i < maxAttempts; i++) {
    const buttons = Array.from(document.querySelectorAll('button')).filter(b => b.offsetParent !== null);
    const btn = buttons.find(b => {
      const text = `${b.textContent} ${b.getAttribute('aria-label')}`.toLowerCase();
      return text.includes('web search') || text.includes('ウェブ検索');
    });
    if (btn) {
      const active = isWebSearchActive(btn);
      console.log('[ensureWebSearchOff] found button', {
        ariaPressed: btn.getAttribute('aria-pressed'),
        color: btn.dataset.colorSchema,
        variant: btn.dataset.variant,
        active,
      });
      if (active) {
        btn.click();
        await sleep(100);
        const stillActive = isWebSearchActive(btn);
        console.log('[ensureWebSearchOff] clicked to disable. success:', !stillActive, { stillActive });
      } else {
        console.log('[ensureWebSearchOff] already inactive');
      }
      return;
    }
    await sleep(interval);
  }
  console.log('[ensureWebSearchOff] toggle not found');
}

// Web検索ボタンを確実にオンにする
export async function ensureWebSearchOn(maxAttempts = 5, interval = 100) {
  for (let i = 0; i < maxAttempts; i++) {
    const buttons = Array.from(document.querySelectorAll('button')).filter(b => b.offsetParent !== null);
    const btn = buttons.find(b => {
      const text = `${b.textContent} ${b.getAttribute('aria-label')}`.toLowerCase();
      return text.includes('web search') || text.includes('ウェブ検索');
    });
    if (btn) {
      const active = isWebSearchActive(btn);
      if (!active) {
        btn.click();
        await sleep(100);
      }
      return;
    }
    await sleep(interval);
  }
  console.log('[ensureWebSearchOn] toggle not found');
}

// Researchボタンを確実にオンにする
export async function ensureResearchOn(maxAttempts = 5, interval = 100) {
  for (let i = 0; i < maxAttempts; i++) {
    const buttons = Array.from(document.querySelectorAll('button')).filter(b => b.offsetParent !== null);
    const btn = buttons.find(b => {
      const text = `${b.textContent} ${b.getAttribute('aria-label')}`.toLowerCase();
      return text.includes('research');
    });
    if (btn) {
      const active = isResearchActive(btn);
      if (!active) {
        btn.click();
        await sleep(100);
      }
      return;
    }
    await sleep(interval);
  }
  console.log('[ensureResearchOn] toggle not found');
}

// Researchボタンを確実にオフにする
export async function ensureResearchOff(maxAttempts = 5, interval = 100) {
  for (let i = 0; i < maxAttempts; i++) {
    const buttons = Array.from(document.querySelectorAll('button')).filter(b => b.offsetParent !== null);
    const btn = buttons.find(b => {
      const text = `${b.textContent} ${b.getAttribute('aria-label')}`.toLowerCase();
      return text.includes('research');
    });
    if (btn) {
      const active = isResearchActive(btn);
      console.log('[ensureResearchOff] found button', {
        ariaPressed: btn.getAttribute('aria-pressed'),
        color: btn.dataset.colorSchema,
        variant: btn.dataset.variant,
        active,
      });
      if (active) {
        btn.click();
        await sleep(100);
        const stillActive = isResearchActive(btn);
        console.log('[ensureResearchOff] clicked to disable. success:', !stillActive, { stillActive });
      } else {
        console.log('[ensureResearchOff] already inactive');
      }
      return;
    }
    await sleep(interval);
  }
}

// 回答の処理が終わるまで待つ
export async function waitForAnswerComplete(timeout = 300000, appearTimeout = 20000) {
  console.log('[waitForAnswerComplete] start', { timeout, appearTimeout });
  const start = Date.now();
  const appearLimit = start + appearTimeout;
  const interval = 200;
  let seen = false;
  while (Date.now() - start < timeout) {
    if (injector.shouldStop) return;
    const spinner = document.querySelector('.animate-pulse');
    if (spinner) {
      if (!seen) {
        console.log('[waitForAnswerComplete] spinner detected');
        seen = true;
      }
    } else if (seen) {
      console.log('[waitForAnswerComplete] spinner disappeared');
      return;
    } else if (Date.now() > appearLimit) {
      console.log('[waitForAnswerComplete] spinner not found');
      return;
    }
    await sleep(interval);
  }
  console.log('[waitForAnswerComplete] timeout');
}

// Researchの読み込み表示が出ている間待機する
async function waitForResearchLoader(timeout = 1200000, appearTimeout = 10000, interval = 1000) {
  console.log('[waitForResearchLoader] start', { timeout, appearTimeout });
  const start = Date.now();
  const appearLimit = start + appearTimeout;
  let found = false;
  while (Date.now() - start < timeout) {
    if (injector.shouldStop) return;
    const loader = document.querySelector(
      '.orbital-loader, .research-loader, [class*="research-loader"], .search-loader'
    );
    if (loader) {
      if (!found) {
        console.log('[waitForResearchLoader] loader detected');
        found = true;
      }
    } else if (found) {
      console.log('[waitForResearchLoader] loader disappeared');
      return;
    } else if (Date.now() > appearLimit) {
      console.log('[waitForResearchLoader] loader not found');
      return;
    }
    await sleep(interval);
  }
  console.log('[waitForResearchLoader] timeout');
}

// Web検索の読み込み表示が出ている間待機する
async function waitForWebSearchLoader(timeout = 1200000, appearTimeout = 10000, interval = 1000) {
  console.log('[waitForWebSearchLoader] start', { timeout, appearTimeout });
  const start = Date.now();
  const appearLimit = start + appearTimeout;
  let found = false;
  while (Date.now() - start < timeout) {
    if (injector.shouldStop) return;
    const loader = document.querySelector(
      '.orbital-loader, .web-search-loader, [class*="web-search-loader"], ' +
        '[class*="search-loader"], [data-testid*="search-loader"], ' +
        '.loading-spinner, svg.animate-spin, .animate-spin, .shimmer-text.animate'
    );
    if (loader) {
      if (!found) {
        console.log('[waitForWebSearchLoader] loader detected');
        found = true;
      }
    } else if (found) {
      console.log('[waitForWebSearchLoader] loader disappeared');
      return;
    } else if (Date.now() > appearLimit) {
      console.log('[waitForWebSearchLoader] loader not found');
      return;
    }
    await sleep(interval);
  }
  console.log('[waitForWebSearchLoader] timeout');
}

// Web検索の処理完了を待つ
async function waitForWebSearchComplete(timeout = 1200000, appearTimeout = 10000, interval = 1000) {
  console.log('[waitForWebSearchComplete] start', { timeout, appearTimeout });
  const start = Date.now();
  const appearLimit = start + appearTimeout;
  let found = false;
  while (Date.now() - start < timeout) {
    if (injector.shouldStop) return;
    const loader = document.querySelector(
      '.orbital-loader, .web-search-loader, [class*="web-search-loader"], ' +
        '[class*="search-loader"], [data-testid*="search-loader"], ' +
        '.loading-spinner, svg.animate-spin, .animate-spin, ' +
        '.shimmer-text.animate'
    );
    if (loader) {
      if (!found) {
        console.log('[waitForWebSearchComplete] loader detected');
        found = true;
      }
    } else if (found) {
      console.log('[waitForWebSearchComplete] loader disappeared');
      return;
    } else if (Date.now() > appearLimit) {
      console.log('[waitForWebSearchComplete] loader not found');
      return;
    }
    await sleep(interval);
  }
  console.log('[waitForWebSearchComplete] timeout');
}

// Deep Research の進行を監視し次のメッセージを送る
export async function monitorDeepResearchLoading(nextMessage, timeout = 1200000, interval = 1000) {
  if ((app.searchMode !== 'deep' && app.searchMode !== 'web') || app.deepMonitorStarted) return;
  app.deepMonitorStarted = true;
  console.log('[monitorDeepResearchLoading] start');
  try {
    if (app.searchMode === 'deep') {
      await waitForResearchLoader(timeout, 10000, interval);
    } else {
      await waitForWebSearchLoader(timeout, 10000, interval);
    }
    await waitForAnswerComplete(300000, 5000);
    if (nextMessage) {
      await ensureResearchOff();
      await ensureWebSearchOff();
      await sendSingleMessage(nextMessage, false, false);
    }
  } finally {
    try {
      await ensureResearchOff();
      await ensureWebSearchOff();
    } catch (e) {
      console.error('toggle off error', e);
    }
    app.deepMonitorStarted = false;
  }
}

// テキストボックスに文字を入力する
async function inputText(el, text) {
  el.focus();
  if (el.value !== undefined) {
    el.value = '';
    await sleep(50);
    el.value = text;
  } else {
    el.textContent = '';
    await sleep(50);
    el.textContent = text;
  }
  el.dispatchEvent(new Event('input', { bubbles: true }));
  await sleep(50);
  el.dispatchEvent(new Event('change', { bubbles: true }));
}

// Enterキーを押したのと同じ動作で送信する
async function sendWithEnter(el) {
  el.focus();
  const ev = new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true });
  el.dispatchEvent(ev);
}

// 送信ボタンをクリックする
async function clickButton() {
  const btn = findElement(injector.config.selectors.button);
  if (!btn) throw new Error('送信ボタンが見つかりません');
  if (btn.disabled) btn.disabled = false;
  btn.click();
}

// 1つのメッセージを送信して結果を待つ
export async function sendSingleMessage(message, watchDeep = true, useResearch) {
  const mode =
    useResearch === true ? 'deep' : useResearch === false ? 'off' : app.searchMode;
  logStep('メッセージ送信処理開始');
  const textbox = findElement(injector.config.selectors.textbox);
  if (!textbox) throw new Error('テキストボックスが見つかりません');

  if (mode === 'deep') {
    await ensureResearchOn();
    await ensureWebSearchOff();
  } else if (mode === 'web') {
    await ensureResearchOff();
    await ensureWebSearchOn();
  } else {
    await ensureResearchOff();
    await ensureWebSearchOff();
  }

  await sleep(200);
  await sleep(injector.config.delays.beforeInput);
  logStep('テキストを入力します');
  await inputText(textbox, message);
  logStep('テキストを入力しました');
  await sleep(injector.config.delays.afterInput);
  if (injector.config.useEnterKey) {
    logStep('送信キーを押します');
    await sendWithEnter(textbox);
  } else {
    logStep('送信ボタンをクリックします');
    await clickButton();
  }
  logStep('送信しました。回答待ち');
  await sleep(injector.config.delays.afterSend);
  if (mode === 'deep') {
    await waitForResearchLoader(300000, 1000);
  } else if (mode === 'web') {
    await waitForWebSearchLoader(300000, 1000);
  } else {
    await waitForWebSearchComplete(500, 500, 100);
  }
  if (watchDeep && app.searchMode !== 'off') {
    const basePrompt = app.multiSlideMode
      ? '上記の調べた以下の手順を実行してください。 ' +
        'トピックからスライドの構成案（セクションやポイント）をまとめます。 ' +
        '各スライドの要点を整理し、スライド順に並べます。 ' +
        'スライドごとのプロンプトをJSON形式で出力 ' +
        'slides 配列に、各スライド用のテキストプロンプトを収めます。 プロンプトは詳細に書いてください。 ' +
        '例: { "slides": [ { "title": "はじめに", "prompt": "トピック概要と重要性を説明するスライド" }, { "title": "メインポイント1", "prompt": "ポイント1の詳細を説明するスライド" } ] } ' +
        '必ずユーザーのOKが出るまで構成案を提案し続けてください。構成案がOKなら「スライド作成開始」ボタンをクリックしてくださいね。' +
        '調べた内容に、何かの仕組みとか構造に関する内容が多い場合は図解をメインとした視覚的にわかりやすい解説スライドにしてください。' +
        '数字のデータが多い場合はグラフをうまく使った視覚的にわかりやすいスライドにしてください。'
      : '上記の調べて内容をもとにスライドを作ってください。' +
        '調べた内容に、何かの仕組みとか構造に関する内容が多い場合は図解をメインとした視覚的にわかりやすい解説スライドにしてください。' +
        '数字のデータが多い場合はグラフをうまく使った視覚的にわかりやすいスライドにしてください。';
    const next = app.slidePrompt ? `${app.slidePrompt} ${basePrompt}` : basePrompt;
    monitorDeepResearchLoading(next).catch(e => console.error(e));
  }
  logStep('回答を受信しました');
}

// 送信予定のメッセージ群を順番に処理する
export async function processMessages() {
  injector.isProcessing = true;
  logStep('メッセージ送信を開始');
  const total = injector.config.messages.length;
  try {
    for (let i = 0; i < total; i++) {
      if (injector.shouldStop) {
        injLog('stopped');
        break;
      }
      const msg = injector.config.messages[i];
      injLog('send', msg);
      await sendSingleMessage(msg);
      await waitForAnswerComplete(300000, 20000);
      logStep('回答を受信しました');
      if (i < total - 1) {
        await sleep(injector.config.delays.betweenMessages);
      }
    }
  } catch (err) {
    throw err;
  } finally {
    injector.isProcessing = false;
    try {
      await ensureResearchOff();
      await ensureWebSearchOff();
    } catch (e) {
      console.error('toggle off error', e);
    }
  }
}

// チャット入力欄に文章を設定して送信する
export function startChat(text, slidePrompt) {
  app.logStart = Date.now();
  logStep('チャットを開始します');
  console.log('[startChat] text:', text);
  console.log('[startChat] slidePrompt:', slidePrompt);
  console.log('[startChat] searchMode:', app.searchMode);
  const data = {
    [app.PREVIEW_PARAM]: '1',
    [app.SEARCH_PARAM]: app.searchMode,
    [app.WAIT_PARAM]: '1',
    logStart: app.logStart,
  };
  if (text) {
    data[app.PROMPT_PARAM] = LZString.compressToEncodedURIComponent(text);
  }
  if (app.searchMode === 'deep') {
    data[app.DEEP_PARAM] = '1';
    if (slidePrompt) {
      data[app.SLIDE_PARAM] = slidePrompt;
    }
  }
  console.log('[startChat] data to save:', data);
  const url = new URL(app.START_CHAT_URL);
  url.searchParams.set(app.WAIT_PARAM, '1');
  if (app.searchMode === 'web') {
    url.searchParams.set(app.SEARCH_PARAM, 'websearch');
  } else if (app.searchMode === 'deep') {
    url.searchParams.set(app.SEARCH_PARAM, 'deepresearch');
  }
  console.log('[startChat] target URL:', url.toString());
  helpers.showWaitOverlay && helpers.showWaitOverlay();
  helpers.handOff && helpers.handOff(data, url.toString());
}

// 複数のプロンプトを順に送信する
export function startPromptList(list, slidePrompt) {
  app.logStart = Date.now();
  logStep('CSVチャットを開始します');
  const total = Array.isArray(list) ? list.length : 0;
  if (total > 0) {
    try {
      helpers.beginCsvFlow && helpers.beginCsvFlow({
        total,
        fileName: app.csvFileName || '',
        resetLogs: true,
      });
    } catch (e) {
      console.error('beginCsvFlow error', e);
    }
  }
  if (Array.isArray(list) && list.length > 1) {
    app.multiSlideMode = true;
  }
  const data = {
    [app.PREVIEW_PARAM]: '1',
    [app.SEARCH_PARAM]: app.searchMode,
    [app.WAIT_PARAM]: '1',
    logStart: app.logStart,
    [app.PROMPT_LIST_PARAM]: list.map(t => LZString.compressToEncodedURIComponent(t)),
  };
  if (app.searchMode === 'deep') {
    data[app.DEEP_PARAM] = '1';
    if (slidePrompt) {
      data[app.SLIDE_PARAM] = slidePrompt;
    }
  }
  const url = new URL(app.CHAT_URL);
  url.searchParams.set(app.WAIT_PARAM, '1');
  if (app.searchMode === 'web') {
    url.searchParams.set(app.SEARCH_PARAM, 'websearch');
  } else if (app.searchMode === 'deep') {
    url.searchParams.set(app.SEARCH_PARAM, 'deepresearch');
  }
  helpers.showWaitOverlay && helpers.showWaitOverlay();
  helpers.handOff && helpers.handOff(data, url.toString());
}

// ページからHTMLコードを取得して自動送信する
export function autoSendPreviewHtml() {
  try {
    const p = app.payload || {};
    if (p[app.PREVIEW_PARAM] !== '1' || !app.lastHtml) return;
    logStep('プレビューHTMLを自動送信します');
    app.searchMode =
      p[app.SEARCH_PARAM] ||
      (p[app.DEEP_PARAM] === '1' ? 'deep' : 'off');
    app.slidePrompt = p[app.SLIDE_PARAM] || '';

    injector.config = {
      messages: [helpers.CONVERT_PREFIX + '\n\n' + app.lastHtml],
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
      debug: false,
    };

    let waited = false;
    injector.timer = setInterval(async () => {
      if (injector.isProcessing) return;
      const box = findElement(injector.config.selectors.textbox);
      if (!box) {
        if (!waited) {
          console.log('[autoSendPreviewHtml] textbox not found');
          waited = true;
        }
        return;
      }
      console.log('[autoSendPreviewHtml] textbox found');
      logStep('テキストボックスを検出しました');
      clearInterval(injector.timer);
      injector.timer = null;
      injector.shouldStop = false;
      helpers.showProgress && helpers.showProgress('creatingPptx', 120000);
      if (app.searchMode === 'deep') {
        await ensureResearchOn();
        await ensureWebSearchOff();
      } else if (app.searchMode === 'web') {
        await ensureResearchOff();
        await ensureWebSearchOn();
      } else {
        await ensureResearchOff();
        await ensureWebSearchOff();
      }
      processMessages()
        .then(() => {
          helpers.updateProgressMessage && helpers.updateProgressMessage('creatingPptx');
          const autoDl = p[app.AUTO_DL_PARAM] === '1';
          helpers.scheduleAutoDownload && helpers.scheduleAutoDownload(0, autoDl);
        })
        .catch(e => console.error(e));
    }, 100);
  } catch (e) {
    console.error('autoSendPreviewHtml', e);
  }
}

// クリップボードなどからプロンプトを読み自動送信する
export function autoSendPrompt() {
  try {
    const p = app.payload || {};
    console.log('[autoSendPrompt] payload:', p);
    console.log('[autoSendPrompt] PROMPT_PARAM:', app.PROMPT_PARAM);
    console.log('[autoSendPrompt] PROMPT_LIST_PARAM:', app.PROMPT_LIST_PARAM);
    const enc = p[app.PROMPT_PARAM];
    const listEnc = p[app.PROMPT_LIST_PARAM];
    console.log('[autoSendPrompt] enc:', enc);
    console.log('[autoSendPrompt] listEnc:', listEnc);
    if (!enc && !Array.isArray(listEnc)) {
      console.log('[autoSendPrompt] No prompt data found, exiting');
      return;
    }
    logStep('プロンプトを自動送信します');
    let messages = [];
    if (Array.isArray(listEnc)) {
      messages = listEnc
        .map(e => {
          try { return LZString.decompressFromEncodedURIComponent(e); } catch { return e; }
        })
        .filter(Boolean);
    } else {
      const text = LZString.decompressFromEncodedURIComponent(enc) || enc;
      messages = [text];
    }
    if (messages.length > 1) app.multiSlideMode = true;
    app.searchMode =
      p[app.SEARCH_PARAM] ||
      (p[app.DEEP_PARAM] === '1' ? 'deep' : 'off');
    app.slidePrompt = p[app.SLIDE_PARAM] || '';

    injector.config = {
      messages: messages,
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
      debug: false,
    };

    let waited = false;
    injector.timer = setInterval(async () => {
      if (injector.isProcessing) return;
      const box = findElement(injector.config.selectors.textbox);
      if (!box) {
        if (!waited) {
          console.log('[autoSendPrompt] textbox not found');
          waited = true;
        }
        return;
      }
      console.log('[autoSendPrompt] textbox found');
      logStep('テキストボックスを検出しました');
      clearInterval(injector.timer);
      injector.timer = null;
      injector.shouldStop = false;
      helpers.monitorHtmlRender && helpers.monitorHtmlRender();
      if (app.searchMode === 'deep') {
        await ensureResearchOn();
        await ensureWebSearchOff();
      } else if (app.searchMode === 'web') {
        await ensureResearchOff();
        await ensureWebSearchOn();
      } else {
        await ensureResearchOff();
        await ensureWebSearchOff();
      }
      processMessages()
        .then(() => {
          if (!injector.shouldStop) {
            helpers.scheduleAutoDownload && helpers.scheduleAutoDownload(0, true);
          }
        })
        .catch(e => console.error(e));
    }, 100);
  } catch (e) {
    console.error('autoSendPrompt', e);
  }
}

// JSON で与えられたスライド情報を送信する
export async function sendSlidesFromJson() {
  try {
    helpers.resetSearchMode && helpers.resetSearchMode(true);

    logStep('スライド送信を開始');
    const slides = await (helpers.extractSlidesJson ? helpers.extractSlidesJson() : []);
    if (!slides || !slides.length) {
      throw new Error('No slides to send');
    }

    helpers.monitorHtmlRender && helpers.monitorHtmlRender();

    if (injector.isProcessing) {
      alert(helpers.t ? helpers.t('processing') : '処理中です');
      return;
    }

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

    injector.shouldStop = false;
    injector.isProcessing = true;

    for (let i = 0; i < slides.length; i++) {
      logStep(`スライド${i + 1}を送信`);
      const slide = slides[i];
      const message = `${helpers.JSON_TO_HTML_PREFIX}---スライド内容---タイトル:${slide.title || ''},プロンプト:${slide.prompt || ''}`;
      await sendSingleMessage(message, false, false);
      await waitForAnswerComplete();
      logStep('回答を受信しました');
      await sleep(injector.config.delays.betweenMessages);
    }
    logStep('スライド送信が完了しました');
  } catch (e) {
    console.error('sendSlidesFromJson error', e);
  } finally {
    injector.isProcessing = false;
  }
}

