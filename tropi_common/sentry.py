"""Fleet-standard Sentry init — the block that was copy-pasted into 12 services.

Usage (top of the entrypoint, before other imports do work):
    from tropi_common.sentry import init_sentry
    init_sentry()

No-op when SENTRY_DSN is unset, so local runs and tests need no config.
Changing fleet policy (sample rate, integrations) = one change here + per-
service pin bumps, instead of 12 PRs.
"""

import os


def init_sentry(**overrides) -> bool:
    """Initialize Sentry from env. Returns True if initialized."""
    dsn = os.getenv("SENTRY_DSN")
    if not dsn:
        return False
    import sentry_sdk

    kwargs = {
        "dsn": dsn,
        "traces_sample_rate": 0.05,
        "environment": os.getenv("RAILWAY_ENVIRONMENT_NAME", "production"),
        "release": os.getenv("RAILWAY_DEPLOYMENT_ID"),
    }
    kwargs.update(overrides)
    sentry_sdk.init(**kwargs)
    return True
