import { next } from "@vercel/functions";

const COOKIE_NAME = "__Host-sapporo_dashboard_session";
const SESSION_TTL_SECONDS = 60 * 60 * 12;
const FAILED_LOGIN_DELAY_MS = 350;
const LOGIN_TITLE = "北海道 就労B型ダッシュボード";
const DASHBOARD_PASSWORD = process.env.DASHBOARD_PASSWORD;
const DASHBOARD_SESSION_SECRET = process.env.DASHBOARD_SESSION_SECRET;

export const config = {
  runtime: "nodejs",
  matcher: ["/((?!_vercel/.*).*)"],
};

export default async function middleware(request) {
  const url = new URL(request.url);

  if (url.pathname === "/logout") {
    return redirectWithCookie("/", clearSessionCookie());
  }

  if (request.method === "POST") {
    return handleLogin(request, url);
  }

  if (await hasValidSession(request)) {
    return next();
  }

  return renderLoginPage(url, { status: 200 });
}

async function handleLogin(request, url) {
  if (!isConfigured()) {
    return renderLoginPage(url, {
      status: 503,
      message: "認証設定がまだ完了していません。",
    });
  }

  let form;
  try {
    form = await request.formData();
  } catch {
    return renderLoginPage(url, {
      status: 400,
      message: "ログインフォームを読み取れませんでした。",
    });
  }
  const password = String(form.get("password") ?? "");
  const nextPath = safeNextPath(form.get("next"), url);

  if (!(await verifyPassword(password))) {
    await wait(FAILED_LOGIN_DELAY_MS);
    return renderLoginPage(url, {
      status: 401,
      message: "パスワードが違います。",
      nextPath,
    });
  }

  return redirectWithCookie(nextPath, await createSessionCookie());
}

async function hasValidSession(request) {
  if (!isConfigured()) return false;
  const cookie = readCookie(request.headers.get("cookie") ?? "", COOKIE_NAME);
  if (!cookie) return false;

  const [payloadValue, signature] = cookie.split(".");
  if (!payloadValue || !signature) return false;

  const payload = decodeBase64UrlToText(payloadValue);
  const [version, expiresAtText] = payload.split(":");
  const expiresAt = Number(expiresAtText);
  if (version !== "v1" || !Number.isFinite(expiresAt) || expiresAt < Date.now()) {
    return false;
  }

  const expectedSignature = await signText(payload);
  return timingSafeEqual(signature, expectedSignature);
}

async function verifyPassword(input) {
  return timingSafeEqual(input, DASHBOARD_PASSWORD ?? "");
}

async function createSessionCookie() {
  const expiresAt = Date.now() + SESSION_TTL_SECONDS * 1000;
  const nonce = crypto.randomUUID();
  const payload = `v1:${expiresAt}:${nonce}`;
  const encodedPayload = encodeTextBase64Url(payload);
  const signature = await signText(payload);
  const expires = new Date(expiresAt).toUTCString();
  return `${COOKIE_NAME}=${encodedPayload}.${signature}; Path=/; Expires=${expires}; HttpOnly; Secure; SameSite=Lax`;
}

function clearSessionCookie() {
  return `${COOKIE_NAME}=; Path=/; Max-Age=0; HttpOnly; Secure; SameSite=Lax`;
}

async function signText(value) {
  const secret = DASHBOARD_SESSION_SECRET ?? "";
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"]
  );
  const signature = await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(value));
  return bytesToBase64Url(new Uint8Array(signature));
}

function isConfigured() {
  return Boolean(DASHBOARD_PASSWORD && DASHBOARD_SESSION_SECRET);
}

function wait(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

function redirectWithCookie(path, cookie) {
  return new Response(null, {
    status: 303,
    headers: {
      "cache-control": "no-store",
      location: path,
      "set-cookie": cookie,
    },
  });
}

function renderLoginPage(url, options = {}) {
  const nextPath = options.nextPath ?? safeNextPath(`${url.pathname}${url.search}${url.hash}`, url);
  const message = options.message ?? "";
  const status = options.status ?? 200;

  return new Response(loginHtml({ message, nextPath }), {
    status,
    headers: {
      "cache-control": "no-store",
      "content-security-policy": "default-src 'self'; style-src 'unsafe-inline'; frame-ancestors 'none'; base-uri 'self'; form-action 'self'",
      "content-type": "text/html; charset=utf-8",
      "permissions-policy": "camera=(), microphone=(), geolocation=()",
      "referrer-policy": "strict-origin-when-cross-origin",
      "x-content-type-options": "nosniff",
      "x-frame-options": "DENY",
    },
  });
}

function loginHtml({ message, nextPath }) {
  const escapedMessage = escapeHtml(message);
  const escapedNext = escapeHtml(nextPath);
  return `<!doctype html>
<html lang="ja">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>${escapeHtml(LOGIN_TITLE)} ログイン</title>
    <style>
      :root {
        color-scheme: light;
        font-family: "Hiragino Sans", "Yu Gothic", sans-serif;
        background: #f6f8fb;
        color: #172033;
      }
      * { box-sizing: border-box; }
      body {
        min-height: 100vh;
        margin: 0;
        display: grid;
        place-items: center;
        padding: 24px;
        background:
          radial-gradient(circle at 15% 10%, rgba(15,118,110,0.18), transparent 30%),
          radial-gradient(circle at 85% 20%, rgba(37,99,235,0.14), transparent 32%),
          linear-gradient(135deg, #f8fafc 0%, #edf5f3 100%);
      }
      main {
        width: min(100%, 420px);
        padding: 28px;
        border: 1px solid rgba(15,23,42,0.08);
        border-radius: 22px;
        background: rgba(255,255,255,0.9);
        box-shadow: 0 22px 60px rgba(15,23,42,0.14);
      }
      .kicker {
        display: inline-flex;
        padding: 5px 10px;
        border-radius: 999px;
        background: #ecfdf5;
        color: #047857;
        font-size: 12px;
        font-weight: 700;
      }
      h1 {
        margin: 16px 0 8px;
        font-size: 24px;
        line-height: 1.35;
      }
      p {
        margin: 0 0 20px;
        color: #64748b;
        font-size: 14px;
        line-height: 1.7;
      }
      label {
        display: grid;
        gap: 8px;
        margin-bottom: 14px;
        color: #334155;
        font-size: 13px;
        font-weight: 700;
      }
      input {
        width: 100%;
        min-height: 46px;
        border: 1px solid #cbd5e1;
        border-radius: 12px;
        padding: 10px 12px;
        font-size: 16px;
        outline: none;
      }
      input:focus {
        border-color: #0f766e;
        box-shadow: 0 0 0 4px rgba(15,118,110,0.12);
      }
      button {
        width: 100%;
        min-height: 46px;
        border: 0;
        border-radius: 12px;
        background: #0f766e;
        color: #fff;
        font-size: 15px;
        font-weight: 800;
        cursor: pointer;
      }
      .message {
        margin: 0 0 14px;
        padding: 10px 12px;
        border-radius: 12px;
        background: #fff7ed;
        color: #9a3412;
        font-size: 13px;
        font-weight: 700;
      }
      .note {
        margin-top: 16px;
        font-size: 12px;
      }
    </style>
  </head>
  <body>
    <main>
      <span class="kicker">Password Required</span>
      <h1>${escapeHtml(LOGIN_TITLE)}</h1>
      <p>パスワードのみで入れる管理用ダッシュボードです。</p>
      ${escapedMessage ? `<div class="message" role="alert">${escapedMessage}</div>` : ""}
      <form method="post">
        <input type="hidden" name="next" value="${escapedNext}" />
        <label>
          パスワード
          <input name="password" type="password" autocomplete="current-password" autofocus required />
        </label>
        <button type="submit">ログイン</button>
      </form>
      <p class="note">ログイン状態は12時間保持されます。共有端末では利用後に /logout を開いてください。</p>
    </main>
  </body>
</html>`;
}

function readCookie(cookieHeader, name) {
  return cookieHeader
    .split(";")
    .map((part) => part.trim())
    .find((part) => part.startsWith(`${name}=`))
    ?.slice(name.length + 1);
}

function safeNextPath(value, currentUrl) {
  const fallback = "/";
  const text = String(value ?? fallback);
  if (!text.startsWith("/") || text.startsWith("//")) return fallback;
  try {
    const nextUrl = new URL(text, currentUrl.origin);
    if (nextUrl.origin !== currentUrl.origin) return fallback;
    return `${nextUrl.pathname}${nextUrl.search}${nextUrl.hash}`;
  } catch {
    return fallback;
  }
}

function encodeTextBase64Url(value) {
  return bytesToBase64Url(new TextEncoder().encode(value));
}

function decodeBase64UrlToText(value) {
  const normalized = value.replaceAll("-", "+").replaceAll("_", "/");
  const padded = normalized.padEnd(normalized.length + ((4 - (normalized.length % 4)) % 4), "=");
  const binary = atob(padded);
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function bytesToBase64Url(bytes) {
  let binary = "";
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary).replaceAll("+", "-").replaceAll("/", "_").replaceAll("=", "");
}

function timingSafeEqual(left, right) {
  const leftText = String(left);
  const rightText = String(right);
  const maxLength = Math.max(leftText.length, rightText.length);
  let diff = leftText.length ^ rightText.length;
  for (let index = 0; index < maxLength; index += 1) {
    diff |= (leftText.charCodeAt(index) || 0) ^ (rightText.charCodeAt(index) || 0);
  }
  return diff === 0;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
