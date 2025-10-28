/**
 * ファイル名: src/apiClient.js
 * 説明:
 *   API経由でパワーポイントを生成するためのクライアントモジュール。
 *   プレビュー/パワポ変換拡張機能のpopup.jsと同じロジックを使用。
 *
 * 主な機能:
 *   - APIキーの保存・読み込み・削除
 *   - API URL設定
 *   - /generate-pptx エンドポイントへのリクエスト送信
 *   - Base64レスポンスをBlobに変換してダウンロード
 *
 * セキュリティ:
 *   - TLS/HTTPS暗号化のみに依存（アプリレベル暗号化は削除）
 *   - APIキーはchrome.storage.localに保存（OSレベルで暗号化）
 *   - IP制限（サーバー側でRakuten INTRA限定）
 */

// API設定
export const API_CONFIG = {
  // APIサーバーのURL
  API_BASE_URL: 'https://powerpoint-genai-test-854259963531.asia-northeast1.run.app',

  // 拡張機能のID（実行時に自動取得）
  // manifest.jsonのkeyフィールドから生成される: mnfcpmjknacajphhdlepejbcbnkllccg
  get EXTENSION_ID() {
    if (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.id) {
      return chrome.runtime.id;
    }
    return 'mnfcpmjknacajphhdlepejbcbnkllccg'; // フォールバック
  },

  // APIキーはchrome.storageから動的に取得
  API_KEY: null,

  // リクエストタイムアウト（5分）
  REQUEST_TIMEOUT: 300000,

  // デバッグモード
  DEBUG: true,
};

// ログ出力ヘルパー
function log(message, ...args) {
  if (API_CONFIG.DEBUG) {
    console.log(`[API Client] ${message}`, ...args);
  }
}

function logError(message, ...args) {
  console.error(`[API Client Error] ${message}`, ...args);
}

/**
 * Background service workerにメッセージを送信
 * @param {Object} message - 送信するメッセージ
 * @returns {Promise<any>} レスポンスデータ
 */
function sendMessageToBackground(message) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(message, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }

      if (!response) {
        reject(new Error('No response from background service worker'));
        return;
      }

      if (!response.success) {
        reject(new Error(response.error?.message || 'Unknown error'));
        return;
      }

      resolve(response.response);
    });
  });
}

/**
 * APIキーをchrome.storage.localから取得
 * @returns {Promise<string|null>} APIキー、または null
 */
export async function getApiKey() {
  return new Promise((resolve) => {
    chrome.storage.local.get(['pptx_api_key'], (result) => {
      // ⚠️ テスト用デフォルトAPIキー（本番環境では削除してください）
      // TODO: 本番デプロイ前に必ずこの行を削除すること
      const DEFAULT_TEST_API_KEY = '9f3ab843049f894c01ab18d2715a92c21522c32803747405f8197b80c82a10a2';

      // ストレージにAPIキーがない場合、デフォルトキーを返す
      const apiKey = result.pptx_api_key || DEFAULT_TEST_API_KEY;

      resolve(apiKey);
    });
  });
}

/**
 * APIキーをchrome.storage.localに保存
 * @param {string} apiKey - 保存するAPIキー
 * @returns {Promise<void>}
 */
export async function saveApiKey(apiKey) {
  return new Promise((resolve) => {
    chrome.storage.local.set({ pptx_api_key: apiKey }, () => {
      log('APIキーを保存しました');
      resolve();
    });
  });
}

/**
 * APIキーをchrome.storage.localから削除
 * @returns {Promise<void>}
 */
export async function clearApiKey() {
  return new Promise((resolve) => {
    chrome.storage.local.remove(['pptx_api_key'], () => {
      log('APIキーをクリアしました');
      resolve();
    });
  });
}

/**
 * API経由でパワーポイントを生成してダウンロード
 *
 * @param {string} script - pptxgenjsコード
 * @param {string} filename - ファイル名（デフォルト: 'presentation.pptx'）
 * @param {Object} payload - 追加データ（オプション）
 * @returns {Promise<void>}
 * @throws {Error} APIキーが未設定、認証エラー、ネットワークエラーなど
 */
export async function generatePptxViaApi(script, filename = 'presentation.pptx', payload = {}) {
  log('API経由でパワーポイント生成を開始');

  // 1. APIキーを確認
  const apiKey = await getApiKey();

  if (!apiKey) {
    throw new Error('APIキーが設定されていません。設定画面からAPIキーを入力してください。');
  }

  log('APIキーを取得しました');

  // 2. リクエストボディを準備
  const requestBody = {
    script: script,
    filename: filename,
    payload: payload
  };

  log('リクエスト送信（TLS暗号化）:', `${API_CONFIG.API_BASE_URL}/generate-pptx`);
  log('Extension ID:', API_CONFIG.EXTENSION_ID);
  log('API Key (first 16 chars):', apiKey.substring(0, 16) + '...');

  // 3. APIリクエスト送信 (background service worker経由)
  const responseData = await sendMessageToBackground({
    action: 'api-fetch',
    url: `${API_CONFIG.API_BASE_URL}/generate-pptx`,
    options: {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-API-Key': apiKey,
        'X-Extension-ID': API_CONFIG.EXTENSION_ID
      },
      body: JSON.stringify(requestBody)
    }
  });

  log('Response received. Status:', responseData.status);

  // 4. エラーハンドリング
  if (responseData.status === 401) {
    throw new Error('APIキーが無効です。設定画面から正しいAPIキーを入力してください。');
  }

  if (responseData.status === 403) {
    const errorData = (typeof responseData.data === 'object') ? responseData.data : {};
    const errorMessage = errorData.message || errorData.error || '';

    logError('403 Forbidden. Error data:', errorData);
    logError('403 Error message:', errorMessage);

    // エラーメッセージからIPアドレスを抽出
    const ipMatch = errorMessage.match(/\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/);
    const currentIp = ipMatch ? ipMatch[0] : null;

    // IP制限のエラーを検出
    if (errorMessage.toLowerCase().includes('ip') ||
        errorMessage.toLowerCase().includes('allowed') ||
        errorMessage.toLowerCase().includes('restricted') ||
        errorMessage.toLowerCase().includes('unauthorized') ||
        errorMessage.toLowerCase().includes('network') ||
        errorMessage.toLowerCase().includes('intra')) {

      const ipInfo = currentIp ? `\n\n現在のIPアドレス: ${currentIp}\n\nこのIPアドレスをサーバー管理者に伝えて、許可リストに追加してもらってください。` : '';

      throw new Error(`🚫 アクセスが拒否されました。

原因: IP制限により、このネットワークからのアクセスは許可されていません。

対処法:
1. VPNを切断してください（Zscaler/Netskope等）
2. Rakuten INTRA社内ネットワーク（Wi-Fi/有線LAN）に直接接続してください
3. または、サーバー管理者に連絡して、現在のIPアドレスを許可リストに追加してもらってください${ipInfo}

エラー詳細: ${errorMessage}`);
    }

    throw new Error(`アクセスが拒否されました。APIキーを確認してください。

エラー詳細: ${errorMessage || '詳細不明'}`);
  }

  if (!responseData.ok) {
    const errorData = (typeof responseData.data === 'object') ? responseData.data : {};
    const errorMessage = errorData.message || errorData.error || `HTTP error! status: ${responseData.status}`;

    log('API Error Response:', errorData);
    throw new Error(errorMessage);
  }

  // 5. レスポンスからBase64データを取得
  const data = (typeof responseData.data === 'object') ? responseData.data : JSON.parse(responseData.data);

  if (!data.data) {
    throw new Error('レスポンスにデータが含まれていません');
  }

  log('レスポンス受信成功');

  // 6. Base64データをBlobに変換
  const byteCharacters = atob(data.data);
  const byteNumbers = new Array(byteCharacters.length);
  for (let i = 0; i < byteCharacters.length; i++) {
    byteNumbers[i] = byteCharacters.charCodeAt(i);
  }
  const byteArray = new Uint8Array(byteNumbers);
  const blob = new Blob([byteArray], {
    type: data.mimeType || 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  });

  log('Blob変換完了:', blob.size, 'bytes');

  // 7. ダウンロード実行
  // ファイル名の処理：サーバーから返されたファイル名を優先し、
  // 文字化けを防ぐためにクライアント側で指定したファイル名を使用
  const finalFilename = filename; // クライアント側で生成したファイル名を使用（日本語対応）

  log('使用するファイル名:', finalFilename);
  log('サーバーから返されたファイル名:', data.filename);

  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = finalFilename; // クライアント側で指定したファイル名を使用
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);

  log('ダウンロード完了:', finalFilename);
}

/**
 * API経由でパワーポイントを生成してBlobを返す（プレビュー用）
 *
 * @param {string} script - pptxgenjsコード
 * @param {string} filename - ファイル名（デフォルト: 'presentation.pptx'）
 * @param {Object} payload - 追加データ（オプション）
 * @returns {Promise<Blob>} 生成されたPPTXファイルのBlob
 * @throws {Error} APIキーが未設定、認証エラー、ネットワークエラーなど
 */
export async function generatePptxBlobViaApi(script, filename = 'presentation.pptx', payload = {}) {
  log('API経由でパワーポイント生成を開始（Blob返却用）');

  // 1. APIキーを確認
  const apiKey = await getApiKey();

  if (!apiKey) {
    throw new Error('APIキーが設定されていません。設定画面からAPIキーを入力してください。');
  }

  log('APIキーを取得しました');

  // 2. リクエストボディを準備
  const requestBody = {
    script: script,
    filename: filename,
    payload: payload
  };

  log('リクエスト送信（TLS暗号化）:', `${API_CONFIG.API_BASE_URL}/generate-pptx`);

  // 3. APIリクエスト送信 (background service worker経由)
  const responseData = await sendMessageToBackground({
    action: 'api-fetch',
    url: `${API_CONFIG.API_BASE_URL}/generate-pptx`,
    options: {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-API-Key': apiKey,
        'X-Extension-ID': API_CONFIG.EXTENSION_ID
      },
      body: JSON.stringify(requestBody)
    }
  });

  log('Response received. Status:', responseData.status);

  // 4. エラーハンドリング
  if (responseData.status === 401) {
    throw new Error('APIキーが無効です。設定画面から正しいAPIキーを入力してください。');
  }

  if (responseData.status === 403) {
    const errorData = (typeof responseData.data === 'object') ? responseData.data : {};
    const errorMessage = errorData.message || errorData.error || '';

    logError('403 Forbidden. Error data:', errorData);

    const ipMatch = errorMessage.match(/\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/);
    const currentIp = ipMatch ? ipMatch[0] : null;

    if (errorMessage.toLowerCase().includes('ip') ||
        errorMessage.toLowerCase().includes('allowed') ||
        errorMessage.toLowerCase().includes('restricted')) {

      const ipInfo = currentIp ? `\n\n現在のIPアドレス: ${currentIp}\n\nこのIPアドレスをサーバー管理者に伝えて、許可リストに追加してもらってください。` : '';

      throw new Error(`🚫 アクセスが拒否されました。

原因: IP制限により、このネットワークからのアクセスは許可されていません。

対処法:
1. VPNを切断してください（Zscaler/Netskope等）
2. Rakuten INTRA社内ネットワーク（Wi-Fi/有線LAN）に直接接続してください
3. または、サーバー管理者に連絡して、現在のIPアドレスを許可リストに追加してもらってください${ipInfo}

エラー詳細: ${errorMessage}`);
    }

    throw new Error(`アクセスが拒否されました。APIキーを確認してください。

エラー詳細: ${errorMessage || '詳細不明'}`);
  }

  if (!responseData.ok) {
    const errorData = (typeof responseData.data === 'object') ? responseData.data : {};
    const errorMessage = errorData.message || errorData.error || `HTTP error! status: ${responseData.status}`;

    log('API Error Response:', errorData);
    throw new Error(errorMessage);
  }

  // 5. レスポンスからBase64データを取得
  const data = (typeof responseData.data === 'object') ? responseData.data : JSON.parse(responseData.data);

  if (!data.data) {
    throw new Error('レスポンスにデータが含まれていません');
  }

  log('レスポンス受信成功');

  // 6. Base64データをBlobに変換
  const byteCharacters = atob(data.data);
  const byteNumbers = new Array(byteCharacters.length);
  for (let i = 0; i < byteCharacters.length; i++) {
    byteNumbers[i] = byteCharacters.charCodeAt(i);
  }
  const byteArray = new Uint8Array(byteNumbers);
  const blob = new Blob([byteArray], {
    type: data.mimeType || 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  });

  log('Blob変換完了:', blob.size, 'bytes');

  return blob;
}

/**
 * API設定の検証
 * @returns {Promise<{valid: boolean, message: string}>}
 */
export async function validateApiConfig() {
  const apiKey = await getApiKey();

  if (!apiKey) {
    return {
      valid: false,
      message: 'APIキーが設定されていません'
    };
  }

  // APIキーの形式チェック（64文字の16進数文字列）
  if (!/^[a-f0-9]{64}$/i.test(apiKey)) {
    return {
      valid: false,
      message: 'APIキーの形式が正しくありません（64文字の16進数である必要があります）'
    };
  }

  return {
    valid: true,
    message: 'API設定は正常です'
  };
}

/**
 * 現在のIPアドレスを取得（IPv4とIPv6の両方）
 * テスト用のダミーリクエストでサーバーに接続し、エラーレスポンスからIPアドレスを取得
 * @returns {Promise<{ipv4: string, ipv6: string, allowed: boolean, message: string}>}
 */
export async function checkCurrentIp() {
  try {
    log('現在のIPアドレスを確認中（サーバー側で取得）...');

    const apiKey = await getApiKey();
    if (!apiKey) {
      throw new Error('APIキーが設定されていません');
    }

    // /generate-pptx に軽量なテストリクエストを送信
    // エラーレスポンスにIPアドレス情報が含まれる
    const testResponse = await sendMessageToBackground({
      action: 'api-fetch',
      url: `${API_CONFIG.API_BASE_URL}/generate-pptx`,
      options: {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-API-Key': apiKey,
          'X-Extension-ID': API_CONFIG.EXTENSION_ID
        },
        body: JSON.stringify({
          script: 'test',
          filename: 'test.pptx',
          payload: {}
        })
      }
    });

    log('IP確認レスポンス:', testResponse);
    log('IP確認レスポンス.ok:', testResponse.ok);
    log('IP確認レスポンス.status:', testResponse.status);

    // ステータスコード200-299の場合は成功（IPは許可されている）
    if (testResponse.ok || (testResponse.status >= 200 && testResponse.status < 300)) {
      // 成功レスポンスから推測
      const responseData = typeof testResponse.data === 'object'
        ? testResponse.data
        : (testResponse.data ? JSON.parse(testResponse.data) : {});

      log('IP確認成功レスポンス:', responseData);
      log('レスポンスヘッダー:', testResponse.headers);

      // レスポンスヘッダーからIPアドレスを取得
      let ipv4 = testResponse.headers?.['x-client-ip'] ||
                 testResponse.headers?.['x-forwarded-for']?.split(',')[0]?.trim() ||
                 testResponse.headers?.['x-real-ip'] ||
                 'Unknown';

      let ipv6 = testResponse.headers?.['x-client-ipv6'] || 'Unknown';

      // IPv4が不明な場合、フォールバック処理
      if (ipv4 === 'Unknown') {
        // リクエストが成功しているので、IPは許可されている
        // ユーザーに分かりやすいメッセージを表示
        ipv4 = '許可済み';
      }

      return {
        ipv4,
        ipv6,
        allowed: true,
        message: 'このIPアドレスはアクセス許可されています',
        details: {}
      };
    }

    // 403エラーの場合はIP制限
    if (testResponse.status === 403) {
      const errorData = typeof testResponse.data === 'object'
        ? testResponse.data
        : {};
      const errorMessage = errorData.message || errorData.error || '';

      log('IP確認403エラー:', errorData);
      log('エラーメッセージ:', errorMessage);

      // エラーメッセージからIPアドレスを抽出（IPv4）
      const ipv4Match = errorMessage.match(/\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/);
      const ipv4 = ipv4Match ? ipv4Match[0] : (errorData.ip || errorData.ipv4 || 'Unknown');

      // エラーメッセージからIPアドレスを抽出（IPv6）
      const ipv6Match = errorMessage.match(/([0-9a-fA-F]{0,4}:){2,7}[0-9a-fA-F]{0,4}/);
      const ipv6 = ipv6Match ? ipv6Match[0] : (errorData.ipv6 || 'Unknown');

      return {
        ipv4,
        ipv6,
        allowed: false,
        message: errorMessage || 'このIPアドレスはアクセス制限されています',
        details: {}
      };
    }

    // その他のエラー（400, 401, 500など）
    const errorData = typeof testResponse.data === 'object'
      ? testResponse.data
      : {};
    const errorMessage = errorData.message || errorData.error || '';

    log('IP確認その他エラー:', errorData);

    // エラーメッセージからIPアドレスを抽出
    const ipv4Match = errorMessage.match(/\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}/);
    const ipv4 = ipv4Match ? ipv4Match[0] : 'Unknown';

    const ipv6Match = errorMessage.match(/([0-9a-fA-F]{0,4}:){2,7}[0-9a-fA-F]{0,4}/);
    const ipv6 = ipv6Match ? ipv6Match[0] : 'Unknown';

    return {
      ipv4,
      ipv6,
      allowed: false,
      message: `確認できませんでした (Status: ${testResponse.status})`,
      details: {}
    };

  } catch (error) {
    logError('IP確認エラー:', error);
    return {
      ipv4: 'Unknown',
      ipv6: 'Unknown',
      allowed: false,
      message: error.message,
      details: {}
    };
  }
}
