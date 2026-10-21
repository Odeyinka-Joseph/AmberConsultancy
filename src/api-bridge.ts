// CORS-free transport for Google Apps Script Web Apps.
// Google Apps Script ContentService does not expose CORS headers for browser fetch.
// This intercepts API POSTs, submits them through a hidden form/iframe, and receives
// the JSON result from the Apps Script HTML response via window.postMessage.

const nativeFetch = window.fetch.bind(window);

type BridgePayload = { __amberBridge?: true; nonce?: string; response?: any };

function isAppsScriptUrl(input: RequestInfo | URL): boolean {
  const url = typeof input === 'string' ? input : input instanceof URL ? input.toString() : input.url;
  return /^https:\/\/script\.google\.com\/macros\/s\/.+\/exec(?:\?.*)?$/i.test(url);
}

function bridgeRequest(url: string, init?: RequestInit): Promise<Response> {
  return new Promise((resolve, reject) => {
    const nonce = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
    const iframeName = `amber_api_${nonce.replace(/[^a-z0-9]/gi, '')}`;
    const iframe = document.createElement('iframe');
    iframe.name = iframeName;
    iframe.style.display = 'none';
    iframe.setAttribute('aria-hidden', 'true');

    const form = document.createElement('form');
    form.method = 'POST';
    form.action = url;
    form.target = iframeName;
    form.style.display = 'none';

    const payload = document.createElement('input');
    payload.type = 'hidden';
    payload.name = 'payload';
    payload.value = typeof init?.body === 'string' ? init.body : JSON.stringify(init?.body ?? {});

    const nonceField = document.createElement('input');
    nonceField.type = 'hidden';
    nonceField.name = 'bridge_nonce';
    nonceField.value = nonce;

    form.append(payload, nonceField);
    document.body.append(iframe, form);

    let settled = false;
    const cleanup = () => {
      window.removeEventListener('message', onMessage);
      form.remove();
      iframe.remove();
    };

    const timer = window.setTimeout(() => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(new TypeError('The Google Sheets service did not respond. Confirm the Apps Script web app is deployed as “Anyone” and contains the AMBER bridge doPost code.'));
    }, 30000);

    const onMessage = (event: MessageEvent<BridgePayload>) => {
      const data = event.data;
      if (event.source !== iframe.contentWindow || !data || data.__amberBridge !== true || data.nonce !== nonce) return;
      settled = true;
      window.clearTimeout(timer);
      cleanup();
      const body = JSON.stringify(data.response ?? { error: 'Empty response from server' });
      resolve(new Response(body, { status: 200, headers: { 'Content-Type': 'application/json' } }));
    };

    window.addEventListener('message', onMessage);
    form.submit();
  });
}

window.fetch = ((input: RequestInfo | URL, init?: RequestInit) => {
  if (isAppsScriptUrl(input) && String(init?.method || 'GET').toUpperCase() === 'POST') {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.toString() : input.url;
    return bridgeRequest(url, init);
  }
  return nativeFetch(input as any, init);
}) as typeof window.fetch;
