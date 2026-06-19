"""Tests for tropi_common.heartbeat — log-and-swallow dead-man's-switch ping."""

import pytest

from tropi_common.heartbeat import ping


class _FakeResp:
    def __init__(self, ok=True, status_code=200, text="OK"):
        self.ok = ok
        self.status_code = status_code
        self.text = text


def test_noop_key_unset(monkeypatch):
    """No-op when HEALTHCHECK_PING_KEY is not set at all."""
    monkeypatch.delenv("HEALTHCHECK_PING_KEY", raising=False)

    calls = []

    def _fake_post(url, **kwargs):
        calls.append(url)
        raise AssertionError("requests.post must not be called when key is unset")

    monkeypatch.setattr("tropi_common.heartbeat.requests.post", _fake_post)

    result = ping("x")
    assert result is False
    assert calls == []


def test_noop_key_empty_string(monkeypatch):
    """No-op when HEALTHCHECK_PING_KEY is an empty string."""
    monkeypatch.setenv("HEALTHCHECK_PING_KEY", "")

    calls = []

    def _fake_post(url, **kwargs):
        calls.append(url)
        raise AssertionError("requests.post must not be called when key is ''")

    monkeypatch.setattr("tropi_common.heartbeat.requests.post", _fake_post)

    result = ping("x")
    assert result is False
    assert calls == []


def test_noop_key_placeholder(monkeypatch):
    """No-op when HEALTHCHECK_PING_KEY is the literal string 'PLACEHOLDER'."""
    monkeypatch.setenv("HEALTHCHECK_PING_KEY", "PLACEHOLDER")

    calls = []

    def _fake_post(url, **kwargs):
        calls.append(url)
        raise AssertionError("requests.post must not be called for PLACEHOLDER key")

    monkeypatch.setattr("tropi_common.heartbeat.requests.post", _fake_post)

    result = ping("x")
    assert result is False
    assert calls == []


def test_success_url_and_body(monkeypatch):
    """Success ping uses the correct URL and posts the detail as bytes."""
    monkeypatch.setenv("HEALTHCHECK_PING_KEY", "abc123")
    monkeypatch.delenv("HEALTHCHECK_BASE_URL", raising=False)

    captured = {}

    def _fake_post(url, data=None, timeout=None):
        captured["url"] = url
        captured["data"] = data
        return _FakeResp(ok=True)

    monkeypatch.setattr("tropi_common.heartbeat.requests.post", _fake_post)

    result = ping("invoice-generator-poll", detail="produced=2")
    assert result is True
    assert captured["url"] == "https://hc-ping.com/abc123/invoice-generator-poll"
    assert captured["data"] == b"produced=2"


def test_failure_url(monkeypatch):
    """Failure ping appends /fail to the slug URL."""
    monkeypatch.setenv("HEALTHCHECK_PING_KEY", "abc123")
    monkeypatch.delenv("HEALTHCHECK_BASE_URL", raising=False)

    captured = {}

    def _fake_post(url, **kwargs):
        captured["url"] = url
        return _FakeResp(ok=True)

    monkeypatch.setattr("tropi_common.heartbeat.requests.post", _fake_post)

    ping("svc", success=False)
    assert captured["url"] == "https://hc-ping.com/abc123/svc/fail"


def test_base_url_override(monkeypatch):
    """HEALTHCHECK_BASE_URL is used instead of the default hc-ping.com."""
    monkeypatch.setenv("HEALTHCHECK_PING_KEY", "abc123")
    monkeypatch.setenv("HEALTHCHECK_BASE_URL", "https://hc.example.com")

    captured = {}

    def _fake_post(url, **kwargs):
        captured["url"] = url
        return _FakeResp(ok=True)

    monkeypatch.setattr("tropi_common.heartbeat.requests.post", _fake_post)

    ping("svc")
    assert captured["url"].startswith("https://hc.example.com/")
    assert captured["url"] == "https://hc.example.com/abc123/svc"


def test_never_raises_on_network_error(monkeypatch):
    """ping() catches exceptions from requests.post and returns False."""
    monkeypatch.setenv("HEALTHCHECK_PING_KEY", "abc123")
    monkeypatch.delenv("HEALTHCHECK_BASE_URL", raising=False)

    def _fake_post(url, **kwargs):
        raise RuntimeError("connection refused")

    monkeypatch.setattr("tropi_common.heartbeat.requests.post", _fake_post)

    result = ping("svc")
    assert result is False
