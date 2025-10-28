/**
 * ファイル名: background.js
 * 説明:
 *   Chrome 拡張のバックグラウンドスクリプト。
 *   コンテンツスクリプトやプレビューパネルからのメッセージを受け取り、
 *   安全な送信元と URL だけが新しいタブを開けるように制御する。
 *
 * 処理の流れ:
 *   1. 受信したメッセージの sender.id とオリジンを検証し、許可リスト外なら遮断。
 *   2. アクションが "open-tab" の場合、開こうとしている URL がホワイトリストに含まれるか確認。
 *   3. 条件を満たしたときのみ chrome.tabs.create でタブを生成し、例外発生時は catch でログを出して安全に失敗する。
 */
// 他のスクリプトから送られてくる命令を受け取り、必要なら新しいタブを開く
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  try {
    // Ensure the message is from this extension and expected origin
    if (sender.id !== chrome.runtime.id) {
      console.warn('Blocked message from unknown sender', sender);
      return;
    }

    const origin = sender.origin || (sender.url ? new URL(sender.url).origin : '');
    const ALLOWED_ORIGINS = [
      'https://r-ai.tsd.public.rakuten-it.com'
    ];
    if (!ALLOWED_ORIGINS.some(o => origin.startsWith(o))) {
      console.warn('Blocked message from unauthorized origin', origin);
      return;
    }

    // Handle API requests
    if (msg && msg.action === 'api-fetch') {
      handleApiFetch(msg, sendResponse);
      return true; // Will respond asynchronously
    }

    if (msg && msg.action === 'open-tab' && typeof msg.url === 'string') {
        const ALLOWED_URLS = [
          'https://r-ai.tsd.public.rakuten-it.com/',
          'https://forms.office.com/',
          'https://chromewebstore.google.com/',
          'https://r10.to/',
        ];

      if (!ALLOWED_URLS.some(u => msg.url.startsWith(u))) {
        console.warn('Blocked open-tab request to non-whitelisted URL:', msg.url);
        return;
      }

      chrome.tabs.create({ url: msg.url, active: !!msg.active });
    }
  } catch (e) {
    console.error('Error processing runtime message', e);
  }
});

// Handle API fetch requests from content scripts
// Note: We cannot force IPv4 at the fetch level because:
// 1. Cloud Run requires SNI (Server Name Indication) for SSL/TLS
// 2. Direct IP access causes certificate validation errors
// 3. Chrome's fetch API doesn't provide low-level network control
//
// Solution: User must disable IPv6 at OS level or API server must support IPv6
async function handleApiFetch(msg, sendResponse) {
  try {
    const { url, options } = msg;

    console.log('[Background] API fetch request:', url);
    console.log('[Background] Request options:', JSON.stringify(options, null, 2));

    const response = await fetch(url, options);

    // Get response data
    const responseData = {
      ok: response.ok,
      status: response.status,
      statusText: response.statusText,
      headers: {}
    };

    // Copy headers
    response.headers.forEach((value, key) => {
      responseData.headers[key] = value;
    });

    // Get body as text first
    const text = await response.text();

    console.log('[Background] Response status:', response.status);
    console.log('[Background] Response text length:', text.length);

    // Try to parse as JSON
    let data;
    try {
      data = JSON.parse(text);
      console.log('[Background] Response data:', JSON.stringify(data, null, 2));
    } catch (e) {
      data = text;
      console.log('[Background] Response as text:', text.substring(0, 200));
    }

    responseData.data = data;

    console.log('[Background] API fetch response status:', responseData.status);
    sendResponse({ success: true, response: responseData });

  } catch (error) {
    console.error('[Background] API fetch error:', error);
    console.error('[Background] Error details:', {
      name: error.name,
      message: error.message,
      stack: error.stack
    });
    sendResponse({
      success: false,
      error: {
        message: error.message,
        name: error.name,
        type: error.constructor.name
      }
    });
  }
}
