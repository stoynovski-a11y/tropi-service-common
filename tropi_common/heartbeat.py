"""Dead-man's-switch ping to healthchecks.io — the fleet liveness signal.

Contract: LOG-AND-SWALLOW, never raises (a watchdog outage must never break a
caller's work cycle — same rule as the telegram module). No-op when
unconfigured, so local runs and tests need no env.

Env:
  HEALTHCHECK_PING_KEY  — project ping key for slug-based pinging. Unset or a
                          placeholder => silent no-op (ships dark, like the
                          POLLER_ENABLED gate in invoice-generator).
  HEALTHCHECK_BASE_URL  — override the ping host (default https://hc-ping.com).

URL model (healthchecks.io slug ping):
  success: {base}/{key}/{slug}
  failure: {base}/{key}/{slug}/fail
"""

import logging
import os

import requests

log = logging.getLogger(__name__)

_PLACEHOLDERS = {"", "PLACEHOLDER"}


def ping(slug: str, *, success: bool = True, detail: str = "") -> bool:
    """Ping the healthchecks.io check identified by `slug`.

    Returns True on a 2xx ping, False otherwise (including the no-op when
    HEALTHCHECK_PING_KEY is unset). Never raises. `detail` becomes the ping
    body (truncated) and shows up in the healthchecks event log for forensics.
    """
    key = os.environ.get("HEALTHCHECK_PING_KEY", "")
    if key in _PLACEHOLDERS:
        return False
    base = os.environ.get("HEALTHCHECK_BASE_URL", "https://hc-ping.com").rstrip("/")
    url = f"{base}/{key}/{slug}" + ("" if success else "/fail")
    try:
        resp = requests.post(url, data=detail[:1000].encode("utf-8"), timeout=10)
        if not resp.ok:
            log.error("healthcheck ping failed: HTTP %s %s",
                      resp.status_code, resp.text[:200])
        return resp.ok
    except Exception as e:
        log.error("healthcheck ping failed: %s", e)
        return False
