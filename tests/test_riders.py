"""Rider modules: sentry init, telegram contract, cc_track fire-and-forget."""

import os
from unittest.mock import MagicMock, patch

from tropi_common import telegram
from tropi_common.sentry import init_sentry


def test_init_sentry_noop_without_dsn():
    with patch.dict(os.environ, {}, clear=False):
        os.environ.pop("SENTRY_DSN", None)
        assert init_sentry() is False


def test_telegram_never_raises():
    with patch.dict(os.environ, {"TELEGRAM_BOT_TOKEN": "t", "TELEGRAM_CHAT_ID": "c"}), \
         patch.object(telegram.requests, "post", side_effect=RuntimeError("api down")):
        assert telegram.send("x") is False  # KA-11 contract: swallow, return False


def test_telegram_noop_when_unconfigured():
    with patch.dict(os.environ, {"TELEGRAM_BOT_TOKEN": "PLACEHOLDER"}):
        assert telegram.send("x") is False


def test_cc_track_import_and_disabled_noop():
    from tropi_common.cc_track import record_flow
    with patch.dict(os.environ, {}, clear=False):
        os.environ.pop("CC_BASE_URL", None)
        record_flow("test-service")  # must not raise without config
