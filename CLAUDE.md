# tropi-service-common — Shared service plumbing for the Тропи Railway fleet

You are in **tropi-service-common**, part of the Тропи Къммодити service fleet. Fleet architecture, the branch-rules table, shared conventions, secrets policy, and the repo index live in [`../SYSTEM_KNOWLEDGE.md`](../SYSTEM_KNOWLEDGE.md). Owner + global rules: `~/CLAUDE.md`.

| | |
|---|---|
| **Platform** | Shared library (**PUBLIC** GitHub repo) |
| **Default branch** | main |
| **Deploy** | pip `git+SHA` pin — consumers pin a commit SHA in their `requirements.txt` |
| **Run locally** | n/a (library; `pip install -e .` for local dev) |

## What it does
Pure-Python shared library providing fleet-wide plumbing so individual services don't duplicate it: safe Excel workbook handling, Sentry initialisation, command-center activity tracking, Telegram notifications, and healthchecks.io dead-man's-switch pinging. All modules are log-and-swallow — they must never crash a caller's work cycle (the "no-raise fleet contract").

## File map
- `tropi_common/excel_safe.py` — `SafeWorkbook` union (keyaccounts base + warehouse-receipts grafts)
- `tropi_common/sentry.py` — `init_sentry()` wrapper
- `tropi_common/cc_track.py` — command-center activity tracking
- `tropi_common/telegram.py` — Telegram `send()` (no-raise)
- `tropi_common/heartbeat.py` — `ping(slug)` → healthchecks.io (no-op when `HEALTHCHECK_PING_KEY` unset)
- `pyproject.toml` — package metadata; requires Python ≥ 3.10

## Gotchas
- **PUBLIC repo** — never commit secrets, tokens, or internal-only paths here.
- Consumers pin an exact SHA (e.g. `@34a479b`). After any change here, bump the pin in each consumer's `requirements.txt` and redeploy.
- `heartbeat.ping()` is a no-op until `HEALTHCHECK_PING_KEY` is set — ships dark safely.
- Most of the fleet pins `c2cedbb`; the heartbeat module landed in `34a479b` (used by the hardened services).
