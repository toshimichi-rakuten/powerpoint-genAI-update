// storage.js
// Storage management utilities separated from previewPanel.js

const secureStorage = (() => {
  const enc = new TextEncoder();
  const dec = new TextDecoder();
  const KEY_NAME = '__secureStorageKey';
  let keyPromise;

  // 利用可能な storage オブジェクトを取得
  function getStorage() {
    if (chrome?.storage?.session) return chrome.storage.session;
    try {
      return sessionStorage;
    } catch {
      return null;
    }
  }

  // 保存されている暗号鍵を読み込む
  async function loadKey() {
    const storage = getStorage();
    if (!storage) return null;
    try {
      if (typeof storage.get === 'function') {
        const res = await storage.get(KEY_NAME);
        return res[KEY_NAME];
      } else {
        return storage.getItem(KEY_NAME);
      }
    } catch {
      return null;
    }
  }

  // 生成した鍵を保存する
  async function saveKey(jwk) {
    const storage = getStorage();
    if (!storage) return;
    if (typeof storage.set === 'function') {
      await storage.set({ [KEY_NAME]: jwk });
    } else {
      storage.setItem(KEY_NAME, jwk);
    }
  }

  // 暗号化に使う鍵を取得（なければ作成）
  async function getKey() {
    if (!keyPromise) {
      keyPromise = (async () => {
        let jwk = await loadKey();
        if (jwk) {
          try {
            return await crypto.subtle.importKey(
              'jwk',
              JSON.parse(jwk),
              { name: 'AES-GCM' },
              false,
              ['encrypt', 'decrypt']
            );
          } catch (e) {
            console.warn('secureStorage key import error', e);
          }
        }
        const key = await crypto.subtle.generateKey(
          { name: 'AES-GCM', length: 256 },
          true,
          ['encrypt', 'decrypt']
        );
        try {
          const exported = await crypto.subtle.exportKey('jwk', key);
          await saveKey(JSON.stringify(exported));
        } catch (e) {
          console.warn('secureStorage key export error', e);
        }
        return key;
      })();
    }
    return keyPromise;
  }

  // 文字列を暗号化する
  async function encrypt(text) {
    const iv = crypto.getRandomValues(new Uint8Array(12));
    const key = await getKey();
    const encrypted = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, enc.encode(text));
    const buffer = new Uint8Array(iv.byteLength + encrypted.byteLength);
    buffer.set(iv, 0);
    buffer.set(new Uint8Array(encrypted), iv.byteLength);
    return btoa(String.fromCharCode(...buffer));
  }

  // 暗号化されたデータを復号する
  async function decrypt(data) {
    const raw = Uint8Array.from(atob(data), (c) => c.charCodeAt(0));
    const iv = raw.slice(0, 12);
    const payload = raw.slice(12);
    const key = await getKey();
    const decrypted = await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, key, payload);
    return dec.decode(decrypted);
  }

  return {
    // 保存された値を復号して取得
    async getItem(k) {
      const storage = getStorage();
      if (!storage) return null;
      try {
        let v;
        if (typeof storage.get === 'function') {
          const res = await storage.get(k);
          v = res[k];
        } else {
          v = storage.getItem(k);
        }
        if (!v) return null;
        return await decrypt(v);
      } catch (e) {
        console.warn('secureStorage decrypt error', e);
        return null;
      }
    },
    // 文字列を暗号化して保存
    async setItem(k, v) {
      const storage = getStorage();
      if (!storage) return;
      try {
        const ev = await encrypt(v);
        if (typeof storage.set === 'function') {
          await storage.set({ [k]: ev });
        } else {
          storage.setItem(k, ev);
        }
      } catch (e) {
        console.warn('secureStorage encrypt error', e);
      }
    },
    // 指定キーのデータを削除
    async removeItem(k) {
      const storage = getStorage();
      if (!storage) return;
      if (typeof storage.remove === 'function') {
        await storage.remove(k);
      } else {
        storage.removeItem(k);
      }
    },
    // 保存済みデータをすべて消去
    async clear() {
      const storage = getStorage();
      if (!storage) return;
      if (typeof storage.clear === 'function') {
        await storage.clear();
      } else {
        storage.clear();
      }
      keyPromise = null;
    },
  };
})();

window.secureStorage = secureStorage;

const HANDOFF_KEY = 'handoffData';
const ANNOUNCEMENT_KEY = 'previewAnnouncementState';

// パラメータ定数
export const PROMPT_PARAM = 'prompt';
export const PROMPT_LIST_PARAM = 'promptList';
export const BULK_TEMPLATES_PARAM = 'bulkTemplates';

// secureStorage から一時データを取り出す
export async function loadPayload() {
  try {
    // sessionStorageを使用（ウェブページから利用可能）
    const data = sessionStorage.getItem(HANDOFF_KEY);
    if (data) {
      sessionStorage.removeItem(HANDOFF_KEY);
      console.log('[loadPayload] Data loaded from sessionStorage');
      return JSON.parse(data);
    }
    console.log('[loadPayload] No data found in sessionStorage');
    return {};
  } catch (e) {
    console.warn('payload retrieval error', e);
    return {};
  }
}

// データを保存してから別ページへ遷移する
export async function handOff(data, target) {
  try {
    const jsonData = JSON.stringify(data);
    console.log('[handOff] Saving data:', data);

    // sessionStorageを使用（ウェブページから利用可能）
    sessionStorage.setItem(HANDOFF_KEY, jsonData);
    console.log('[handOff] Data saved to sessionStorage');
  } catch (e) {
    console.error('handOff error', e);
  }
  location.href = target;
}

function parseAnnouncementState(raw) {
  if (!raw) return {};
  if (typeof raw === 'string') {
    try {
      return JSON.parse(raw) || {};
    } catch (e) {
      console.warn('announcement state parse error', e);
      return {};
    }
  }
  if (typeof raw === 'object') return raw;
  return {};
}

export function loadAnnouncementState(app, cb) {
  const assignState = (state) => {
    const parsed = parseAnnouncementState(state);
    app.announcementState = parsed;
    if (cb) cb(parsed);
  };

  if (chrome && chrome.storage && chrome.storage.local) {
    chrome.storage.local.get(ANNOUNCEMENT_KEY, (res = {}) => {
      assignState(res[ANNOUNCEMENT_KEY]);
    });
  } else {
    const raw = localStorage.getItem(ANNOUNCEMENT_KEY);
    assignState(raw);
  }
}

export function saveAnnouncementState(state) {
  const value = state || {};
  if (chrome && chrome.storage && chrome.storage.local) {
    chrome.storage.local.set({ [ANNOUNCEMENT_KEY]: value });
  } else {
    try {
      localStorage.setItem(ANNOUNCEMENT_KEY, JSON.stringify(value));
    } catch (e) {
      console.warn('announcement state save error', e);
    }
  }
}

// 登録状態を読み込みアプリに設定する
export function loadRegStatus(app, cb) {
  // ========================================
  // テスト用: 初回登録フラグをデフォルトでOFF（登録済み状態）
  // 本番に戻すときは以下の2行をコメントアウトして、下のコードのコメントを外す
  app.previewRegDone = true;
  app.pptxRegDone = true;
  if (cb) cb();
  return;
  // ========================================

  /* 本番用コード（上のreturnをコメントアウトしてこちらを有効化）
  if (chrome && chrome.storage && chrome.storage.local) {
    chrome.storage.local.get(
      ['previewRegDone', 'pptxRegDone', 'searchMode'],
      (res = {}) => {
        app.previewRegDone = !!res.previewRegDone;
        app.pptxRegDone = !!res.pptxRegDone;
        if (res.searchMode) app.searchMode = res.searchMode;
        if (cb) cb();
      }
    );
  } else {
    app.previewRegDone = localStorage.getItem('previewRegDone') === '1';
    app.pptxRegDone = localStorage.getItem('pptxRegDone') === '1';
    const search = localStorage.getItem('searchMode');
    if (search) app.searchMode = search;
    if (cb) cb();
  }
  */
}

// 登録状態を保存する
export function saveRegStatus(app) {
  if (chrome && chrome.storage && chrome.storage.local) {
    chrome.storage.local.set({
      previewRegDone: app.previewRegDone,
      pptxRegDone: app.pptxRegDone,
      searchMode: app.searchMode,
    });
  } else {
    localStorage.setItem('previewRegDone', app.previewRegDone ? '1' : '');
    localStorage.setItem('pptxRegDone', app.pptxRegDone ? '1' : '');
    localStorage.setItem('searchMode', app.searchMode);
  }
}

// パネルの開閉状態を読み込む
export function loadPanelState(app, cb) {
  if (chrome && chrome.storage && chrome.storage.local) {
    chrome.storage.local.get('panelOpen', (res = {}) => {
      app.panelOpen = res.panelOpen === '1';
      if (cb) cb();
    });
  } else {
    app.panelOpen = localStorage.getItem('panelOpen') === '1';
    if (cb) cb();
  }
}

// パネルの開閉状態を保存する
export function savePanelState(app) {
  if (chrome && chrome.storage && chrome.storage.local) {
    chrome.storage.local.set({ panelOpen: app.panelOpen ? '1' : '' });
  } else {
    localStorage.setItem('panelOpen', app.panelOpen ? '1' : '');
  }
}

export { secureStorage };

// ========================================
// API経由ダウンロード機能のためのAPIキー管理
// ========================================

/**
 * APIキーをchrome.storage.localから取得
 * @returns {Promise<string|null>} APIキー、または null
 */
export async function loadApiKey() {
  return new Promise((resolve) => {
    if (chrome && chrome.storage && chrome.storage.local) {
      chrome.storage.local.get(['pptx_api_key'], (result) => {
        resolve(result.pptx_api_key || null);
      });
    } else {
      try {
        const key = localStorage.getItem('pptx_api_key');
        resolve(key || null);
      } catch {
        resolve(null);
      }
    }
  });
}

/**
 * APIキーをchrome.storage.localに保存
 * @param {string} apiKey - 保存するAPIキー
 * @returns {Promise<void>}
 */
export async function saveApiKey(apiKey) {
  return new Promise((resolve) => {
    if (chrome && chrome.storage && chrome.storage.local) {
      chrome.storage.local.set({ pptx_api_key: apiKey }, () => {
        console.log('[Storage] APIキーを保存しました');
        resolve();
      });
    } else {
      try {
        localStorage.setItem('pptx_api_key', apiKey);
        console.log('[Storage] APIキーを保存しました');
        resolve();
      } catch (e) {
        console.error('[Storage] APIキーの保存に失敗しました', e);
        resolve();
      }
    }
  });
}

/**
 * APIキーをchrome.storage.localから削除
 * @returns {Promise<void>}
 */
export async function clearApiKey() {
  return new Promise((resolve) => {
    if (chrome && chrome.storage && chrome.storage.local) {
      chrome.storage.local.remove(['pptx_api_key'], () => {
        console.log('[Storage] APIキーをクリアしました');
        resolve();
      });
    } else {
      try {
        localStorage.removeItem('pptx_api_key');
        console.log('[Storage] APIキーをクリアしました');
        resolve();
      } catch (e) {
        console.error('[Storage] APIキーのクリアに失敗しました', e);
        resolve();
      }
    }
  });
}
