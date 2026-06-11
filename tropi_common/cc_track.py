"""
cc_track.py — Fire-and-forget time savings tracker for Command Center.

Usage:
    from tropi_common.cc_track import record_flow
    record_flow("svedenie")               # 1 run, 1 item
    record_flow("invoice-renamer", items=5)
    record_flow("bank-parser", meta={"file": "statement.xml"})

Shipped via tropi-service-common (was: copied into each service). Env vars:
    CC_BASE_URL  — e.g. https://command-center-beta-fawn.vercel.app
    CC_TOKEN     — same value as CRON_SECRET on the Command Center
"""

import os
import threading
import logging

log = logging.getLogger(__name__)


def record_flow(flow_slug: str, items: int = 1, meta: dict | None = None) -> None:
    """Send a flow run record to the Command Center (non-blocking)."""
    base_url = os.environ.get("CC_BASE_URL", "").rstrip("/")
    token = os.environ.get("CC_TOKEN", "")
    if not base_url or not token:
        return  # silently skip if not configured

    def _send():
        try:
            import urllib.request, json
            payload = json.dumps({"flow_slug": flow_slug, "items_count": items, "meta": meta}).encode()
            req = urllib.request.Request(
                f"{base_url}/api/flows/record",
                data=payload,
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {token}",
                },
                method="POST",
            )
            with urllib.request.urlopen(req, timeout=5) as resp:
                if resp.status not in (200, 201):
                    log.debug("cc_track: unexpected status %s", resp.status)
        except Exception as e:
            log.debug("cc_track: failed to record flow %s: %s", flow_slug, e)

    threading.Thread(target=_send, daemon=True).start()
