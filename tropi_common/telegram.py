"""Fleet-standard Telegram sender — the skeleton three services re-implemented.

Contract: LOG-AND-SWALLOW, never raises (audit KA-11: a raising sender let a
Telegram outage abort an intake poll cycle). Per-service message BUILDERS
stay local — this module only sends.

Env: TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID (default chat; per-call override
via chat_id=). Unset/placeholder token = silent no-op (sales-autofill runs
with notifications deliberately disabled).
"""

import logging
import os

import requests

log = logging.getLogger(__name__)

_PLACEHOLDERS = {"", "PLACEHOLDER"}


def _is_configured() -> bool:
    return os.environ.get("TELEGRAM_BOT_TOKEN", "") not in _PLACEHOLDERS


def send(text: str, chat_id: str | None = None, parse_mode: str = "HTML") -> bool:
    """Send a message. Returns True on success, False otherwise. Never raises."""
    if not _is_configured():
        return False
    token = os.environ["TELEGRAM_BOT_TOKEN"]
    chat = chat_id or os.environ.get("TELEGRAM_CHAT_ID", "")
    if chat in _PLACEHOLDERS:
        return False
    try:
        resp = requests.post(
            f"https://api.telegram.org/bot{token}/sendMessage",
            json={"chat_id": chat, "text": text, "parse_mode": parse_mode},
            timeout=15,
        )
        if not resp.ok:
            log.error("Telegram send failed: HTTP %s %s", resp.status_code, resp.text[:200])
        return resp.ok
    except Exception as e:
        log.error("Telegram send failed: %s", e)
        return False
