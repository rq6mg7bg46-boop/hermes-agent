"""Regression test: QQBot must NOT go permanently dark on transport failures.

Context
--------
On 2026-10-24 the WSL DNS proxy (10.255.255.254) blackholed the Tencent
hosts for ~5 hours.  QQBot's reconnect loop gave up after
MAX_RECONNECT_ATTEMPTS (100) and called _mark_disconnected(), so the adapter
stayed dead for the rest of the gateway lifetime while Feishu / Weixin kept
retrying and recovered on their own.

These tests pin the fix: transport-level failures (DNS / network) reset the
backoff counter and keep retrying forever instead of disconnecting.
"""

import asyncio

import pytest

from gateway.config import PlatformConfig
from gateway.platforms.qqbot.adapter import QQAdapter
from gateway.platforms.qqbot.constants import (
    MAX_RECONNECT_ATTEMPTS,
    RECONNECT_BACKOFF,
    RETRY_FOREVER_ON_TRANSPORT_ERROR,
)


def _make_adapter(monkeypatch, reconnect_results):
    """Build a QQAdapter whose _reconnect() returns canned results."""
    adapter = QQAdapter(PlatformConfig(enabled=True, extra={}))
    adapter._running = True
    adapter._ws = None
    adapter._session_id = "sess"
    adapter._last_seq = None
    adapter._heartbeat_interval = 30.0
    adapter._access_token = "tok"
    adapter._token_expires_at = 0.0

    calls = {"n": 0, "max_idx": 0}

    async def fake_reconnect(backoff_idx):
        calls["n"] += 1
        calls["max_idx"] = max(calls["max_idx"], backoff_idx)
        if reconnect_results:
            return reconnect_results.pop(0)
        return False

    monkeypatch.setattr(adapter, "_reconnect", fake_reconnect)
    monkeypatch.setattr(adapter, "_mark_transport_disconnected", lambda: None)
    monkeypatch.setattr(adapter, "_fail_pending", lambda reason: None)

    return adapter, calls


async def _run_listen_loop(adapter, monkeypatch, iterations):
    """Drive QQAdapter._listen_loop's transport-error branch.

    _listen_loop is an infinite `while self._running` loop.  We stop it by
    flipping _running off after enough iterations.  Every iteration raises
    inside the try block so the generic `except Exception` branch (the
    transport-error path) is the one under test.
    """
    state = {"disconnected": False, "iterations": 0}

    original_mark = adapter._mark_disconnected

    def tracking_mark_disconnected():
        state["disconnected"] = True
        original_mark()

    monkeypatch.setattr(adapter, "_mark_disconnected", tracking_mark_disconnected)

    async def fake_read_events():
        state["iterations"] += 1
        if state["iterations"] >= iterations:
            adapter._running = False
        raise RuntimeError(
            "Blocked request - DNS resolution failed for: api.sgroup.qq.com"
        )

    monkeypatch.setattr(adapter, "_read_events", fake_read_events)

    await asyncio.wait_for(adapter._listen_loop(), timeout=30)
    return state


@pytest.mark.asyncio
async def test_transport_failure_beyond_attempt_cap_does_not_disconnect(monkeypatch):
    """The adapter must not call _mark_disconnected() on transport errors."""
    assert RETRY_FOREVER_ON_TRANSPORT_ERROR is True

    adapter, calls = _make_adapter(monkeypatch, [])

    state = await _run_listen_loop(
        adapter, monkeypatch, MAX_RECONNECT_ATTEMPTS + 30
    )

    assert state["disconnected"] is False, (
        "QQBot went permanently dark after transport failures — this is the "
        "2026-10-24 outage regression"
    )
    # It should still be actively retrying well past the old cap.
    assert calls["n"] >= MAX_RECONNECT_ATTEMPTS


@pytest.mark.asyncio
async def test_backoff_index_never_exceeds_cap_during_long_outage(monkeypatch):
    """The reconnect delay stays capped even across a very long outage."""
    adapter, calls = _make_adapter(monkeypatch, [])

    await _run_listen_loop(adapter, monkeypatch, MAX_RECONNECT_ATTEMPTS + 30)

    assert calls["max_idx"] <= MAX_RECONNECT_ATTEMPTS
    # RECONNECT_BACKOFF is indexed by backoff_idx, so the sleep never exceeds
    # the last entry (60s) — the retry loop stays cheap.
    assert RECONNECT_BACKOFF[-1] == 60


@pytest.mark.asyncio
async def test_successful_reconnect_after_long_outage_recovers(monkeypatch):
    """After the outage clears, the adapter reconnects instead of staying dead."""
    adapter, calls = _make_adapter(
        monkeypatch, [False] * (MAX_RECONNECT_ATTEMPTS + 5) + [True]
    )

    state = await _run_listen_loop(adapter, monkeypatch, MAX_RECONNECT_ATTEMPTS + 20)

    assert state["disconnected"] is False
    assert calls["n"] > MAX_RECONNECT_ATTEMPTS


def test_retry_forever_flag_is_enabled():
    """Guard against the flag being flipped off by an upgrade."""
    assert RETRY_FOREVER_ON_TRANSPORT_ERROR is True
