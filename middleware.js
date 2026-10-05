// Vercel Routing Middleware: sign-in gate for the whole site.
// Credentials come from Vercel environment variables, never from this (public) repo:
//   LOGIN_EMAIL     e.g. the authorised user's email (compared case-insensitively)
//   LOGIN_PASSWORD  the password
//   AUTH_SECRET     optional, a long random string used to sign session cookies
// If LOGIN_EMAIL or LOGIN_PASSWORD is missing, every page returns 503 (fails closed).

export const config = {
  // everything except the login page's own logo and the favicon
  matcher: '/((?!joulewise-logo\\.jpg|favicon\\.ico).*)',
};

const COOKIE = 'jw_session';
const MAX_AGE = 60 * 60 * 12; // 12 hours
const enc = new TextEncoder();

const b64url = (bytes) => {
  let s = '';
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
};
const fromB64url = (s) => atob(s.replace(/-/g, '+').replace(/_/g, '/'));

async function sign(secret, msg) {
  const key = await crypto.subtle.importKey('raw', enc.encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']);
  return b64url(new Uint8Array(await crypto.subtle.sign('HMAC', key, enc.encode(msg))));
}

function sameString(a, b) {
  if (a.length !== b.length) return false;
  let r = 0;
  for (let i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}

function settings() {
  const email = (process.env.LOGIN_EMAIL || '').trim().toLowerCase();
  const password = process.env.LOGIN_PASSWORD || '';
  const secret = process.env.AUTH_SECRET || `${email}:${password}:jw`;
  return { email, password, secret, ok: Boolean(email && password) };
}

async function makeToken(s) {
  const payload = `${s.email}|${Math.floor(Date.now() / 1000) + MAX_AGE}`;
  return `${b64url(enc.encode(payload))}.${await sign(s.secret, payload)}`;
}

async function validToken(s, token) {
  if (!token || !token.includes('.')) return false;
  const [p, sig] = token.split('.');
  let payload;
  try { payload = fromB64url(p); } catch { return false; }
  const [email, exp] = payload.split('|');
  if (email !== s.email || !(Number(exp) > Date.now() / 1000)) return false;
  return sameString(sig, await sign(s.secret, payload));
}

function readCookie(request, name) {
  const raw = request.headers.get('cookie') || '';
  for (const part of raw.split(';')) {
    const [k, ...v] = part.trim().split('=');
    if (k === name) return v.join('=');
  }
  return '';
}

// only same-site paths, never "//host" or absolute URLs
const safeNext = (n) => (typeof n === 'string' && n.startsWith('/') && !n.startsWith('//') && !n.startsWith('/login') ? n : '/');

const redirect = (url, extraHeaders = {}) => new Response(null, { status: 303, headers: { Location: url, 'Cache-Control': 'no-store', ...extraHeaders } });

export default async function middleware(request) {
  const url = new URL(request.url);
  const s = settings();

  if (!s.ok) {
    return new Response('Sign-in is not configured. Set LOGIN_EMAIL and LOGIN_PASSWORD in the Vercel project settings, then redeploy.', {
      status: 503, headers: { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store' },
    });
  }

  const secure = url.protocol === 'https:' ? '; Secure' : '';

  if (url.pathname === '/logout') {
    return redirect('/login', { 'Set-Cookie': `${COOKIE}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax${secure}` });
  }

  if (url.pathname === '/login' || url.pathname === '/login.html') {
    if (request.method === 'POST') {
      let email = '', password = '', next = '/';
      try {
        const form = await request.formData();
        email = String(form.get('email') || '').trim().toLowerCase();
        password = String(form.get('password') || '');
        next = safeNext(String(form.get('next') || '/'));
      } catch { /* malformed form: treated as a failed sign-in */ }
      if (sameString(email, s.email) && sameString(password, s.password)) {
        return redirect(next, { 'Set-Cookie': `${COOKIE}=${await makeToken(s)}; Path=/; Max-Age=${MAX_AGE}; HttpOnly; SameSite=Lax${secure}` });
      }
      return redirect(`/login?e=1&next=${encodeURIComponent(next)}`);
    }
    // already signed in: skip the form
    if (await validToken(s, readCookie(request, COOKIE))) return redirect(safeNext(url.searchParams.get('next') || '/'));
    return; // serve the static login page
  }

  if (await validToken(s, readCookie(request, COOKIE))) return; // signed in: continue to the page

  return redirect(`/login?next=${encodeURIComponent(url.pathname + url.search)}`);
}
