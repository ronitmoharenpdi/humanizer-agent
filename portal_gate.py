"""
portal_gate.py — only let people in who came through the NP Digital AI Portal.

Drop this file next to app.py in any FastAPI / Starlette agent and add:

    from portal_gate import PortalGateMiddleware
    app.add_middleware(PortalGateMiddleware)

Environment variables (set them in Vercel → the agent's project → Settings):

    AGENT_LAUNCH_SECRET  same long random value as in the portal. While it is
                         unset the gate stays OPEN (no lock), so the agent
                         keeps working during setup and local development.
    AGENT_SLUG           this agent's slug in the portal, e.g. content-outline-generator
    PORTAL_URL           e.g. https://your-portal.vercel.app (where to send strangers)

How it works:
    1. In the portal, "Open agent" goes to /launch/<slug>. The portal checks
       the login and redirects here with ?portal_token=<signed, 2-minute token>.
    2. This middleware verifies the token, sets an 8-hour signed cookie and
       reloads the page without the token in the address bar.
    3. Anyone opening the agent's URL directly (copied link, bookmark after
       the cookie expired, etc.) gets a "please open it from the portal" page.

Uses only the Python standard library.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import time
from http.cookies import SimpleCookie
from urllib.parse import parse_qsl, urlencode

TOKEN_PARAM = "portal_token"
COOKIE_NAME = "np_portal_session"
SESSION_SECONDS = 8 * 60 * 60
ALLOWED_DOMAIN = "neilpateldigital.in"

# Paths that stay open (health checks and static assets carry no data).
OPEN_PATH_PREFIXES = ("/health", "/static/", "/favicon.ico")


def _b64url_decode(value: str) -> bytes:
    return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))


def _b64url_encode(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).rstrip(b"=").decode()


def _sign(payload: str, secret: str) -> str:
    return _b64url_encode(hmac.new(secret.encode(), payload.encode(), hashlib.sha256).digest())


def _verify(token: str, secret: str) -> dict | None:
    """Return the payload if the signature is valid and it hasn't expired."""
    try:
        payload_b64, signature = token.split(".", 1)
    except ValueError:
        return None
    if not hmac.compare_digest(_sign(payload_b64, secret), signature):
        return None
    try:
        payload = json.loads(_b64url_decode(payload_b64))
    except (ValueError, json.JSONDecodeError):
        return None
    if not isinstance(payload, dict) or int(payload.get("exp", 0)) < time.time():
        return None
    email = str(payload.get("sub", "")).lower()
    if not email.endswith("@" + ALLOWED_DOMAIN):
        return None
    return payload


def _make_session(email: str, secret: str) -> str:
    payload = _b64url_encode(
        json.dumps({"sub": email, "exp": int(time.time()) + SESSION_SECONDS}).encode()
    )
    return f"{payload}.{_sign(payload, secret)}"


class PortalGateMiddleware:
    """Pure ASGI middleware, so streaming responses keep working."""

    def __init__(self, app):
        self.app = app
        self.secret = os.environ.get("AGENT_LAUNCH_SECRET", "").strip()
        if not self.secret:
            print("[portal_gate] AGENT_LAUNCH_SECRET not set - gate is OPEN (no lock).")
        self.slug = os.environ.get("AGENT_SLUG", "")
        self.portal_url = os.environ.get("PORTAL_URL", "").rstrip("/")

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            return await self.app(scope, receive, send)

        path = scope.get("path", "/")
        if path == "/health" or path.startswith(OPEN_PATH_PREFIXES):
            return await self.app(scope, receive, send)

        # Not configured yet (no AGENT_LAUNCH_SECRET): stay open so the agent
        # keeps working while the secret is being set up, e.g. locally.
        # Set AGENT_LAUNCH_SECRET in Vercel to switch the lock on.
        if not self.secret:
            return await self.app(scope, receive, send)

        query = dict(parse_qsl(scope.get("query_string", b"").decode()))
        token = query.pop(TOKEN_PARAM, None)

        # 1) Fresh launch from the portal → set cookie, strip token from URL.
        if token:
            payload = _verify(token, self.secret)
            if payload and (not self.slug or payload.get("aud") == self.slug):
                clean = path + (f"?{urlencode(query)}" if query else "")
                cookie = (
                    f"{COOKIE_NAME}={_make_session(payload['sub'], self.secret)}; "
                    f"Max-Age={SESSION_SECONDS}; Path=/; HttpOnly; Secure; SameSite=Lax"
                )
                await send({
                    "type": "http.response.start",
                    "status": 303,
                    "headers": [
                        (b"location", clean.encode()),
                        (b"set-cookie", cookie.encode()),
                        (b"cache-control", b"no-store"),
                    ],
                })
                return await send({"type": "http.response.body", "body": b""})
            return await self._deny(scope, send, "This launch link has expired. Please open the agent again from the portal.")

        # 2) Returning visitor with a valid session cookie.
        cookies = SimpleCookie()
        for name, value in scope.get("headers", []):
            if name == b"cookie":
                cookies.load(value.decode("latin-1"))
        session = cookies.get(COOKIE_NAME)
        if session and _verify(session.value, self.secret):
            return await self.app(scope, receive, send)

        # 3) Everyone else.
        return await self._deny(scope, send, "Please open this agent from the NP Digital AI Portal.")

    async def _deny(self, scope, send, message: str):
        accepts_html = any(
            name == b"accept" and b"text/html" in value for name, value in scope.get("headers", [])
        )
        if accepts_html:
            link = (
                f'<a href="{self.portal_url}/dashboard" style="display:inline-block;margin-top:20px;'
                'padding:10px 18px;border-radius:8px;background:#FF5A1F;color:#fff;'
                'text-decoration:none;font-weight:600">Go to the portal</a>'
                if self.portal_url
                else ""
            )
            body = (
                "<!doctype html><html><head><meta charset='utf-8'>"
                "<meta name='viewport' content='width=device-width,initial-scale=1'>"
                "<title>Access via portal</title></head>"
                "<body style='margin:0;min-height:100vh;display:grid;place-items:center;"
                "background:#0B0D12;color:#F5F7FA;font-family:system-ui,sans-serif;text-align:center'>"
                "<div style='padding:24px'><h1 style='font-size:22px;margin:0 0 8px'>Access restricted</h1>"
                f"<p style='color:#A3ADBD;margin:0'>{message}</p>{link}</div></body></html>"
            ).encode()
            content_type = b"text/html; charset=utf-8"
        else:
            body = json.dumps({"success": False, "error": message}).encode()
            content_type = b"application/json"

        await send({
            "type": "http.response.start",
            "status": 401,
            "headers": [(b"content-type", content_type), (b"cache-control", b"no-store")],
        })
        await send({"type": "http.response.body", "body": body})
