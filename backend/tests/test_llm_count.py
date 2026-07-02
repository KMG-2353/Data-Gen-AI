"""Exact test-case-count guarantee for generate_test_data (batching).

Defect class: on a large single-shot request the LLM truncates its JSON — asked
for 50 rows it returns ~16 ("generating 50, it's generating for 16"). The fix
requests rows in small batches and loops until the FULL count is reached, then
trims/pads to exactly the requested number for ANY N (100/200/500 …).

These tests mock the OpenAI client so no network/key is needed.
"""
import json
import re

import pytest

from app import llm_service


class _Msg:
    def __init__(self, content):
        self.message = type("_M", (), {"content": content})()


class _Resp:
    def __init__(self, content):
        self.choices = [_Msg(content)]


class _Completions:
    """Fake chat.completions honouring the per-call 'Generate exactly N rows'.

    ``mode`` controls how many rows the fake returns relative to the request:
      * "exact" — returns exactly N (tests the batching loop hits the total).
      * "over"  — returns N+extra    (tests trimming to the requested total).
      * "under" — returns a fraction (tests padding up to the requested total).
    """

    def __init__(self, mode: str):
        self.mode = mode
        self.calls: list[str] = []

    def create(self, *, model, temperature, response_format, messages):
        prompt = messages[-1]["content"]
        self.calls.append(prompt)
        n = int(re.search(r"Generate exactly (\d+) rows", prompt).group(1))
        if self.mode == "exact":
            produce = n
        elif self.mode == "over":
            produce = n + 3
        else:  # under
            produce = max(1, n // 5)
        call_ix = len(self.calls)
        rows = [
            {"Test ID": f"T{call_ix}-{i}", "Name": f"Name-{call_ix}-{i}"}
            for i in range(produce)
        ]
        return _Resp(json.dumps({"data": rows}))


class _Client:
    def __init__(self, mode: str):
        self.chat = type("_Chat", (), {"completions": _Completions(mode)})()


@pytest.fixture
def patch_client(monkeypatch):
    def _install(mode: str) -> _Client:
        client = _Client(mode)
        monkeypatch.setattr(llm_service, "configure_openai", lambda: client)
        return client

    return _install


HEADERS = ["Test ID", "Name"]


def _gen(n: int):
    return llm_service.generate_test_data(
        headers=HEADERS, row_count=n, special_instruction="", sheet_name="Policy Info"
    )


@pytest.mark.parametrize("n", [16, 50, 100, 200, 500])
def test_exact_count_any_number(patch_client, n):
    # The headline defect: any requested N must come back as exactly N rows.
    patch_client("exact")
    out = _gen(n)
    assert len(out) == n
    assert all(set(r.keys()) == set(HEADERS) for r in out)


def test_batches_until_full_and_continues(patch_client):
    # 50 rows @ batch 20 → 3 calls (20 + 20 + 10); later calls are continuations.
    client = patch_client("exact")
    out = _gen(50)
    assert len(out) == 50
    calls = client.chat.completions.calls
    assert len(calls) == 3
    # First call is a fresh request; subsequent calls carry the continuation hint
    # so the model produces distinct, non-duplicate rows.
    assert "BATCH CONTINUATION" not in calls[0]
    assert "BATCH CONTINUATION" in calls[1]
    assert "BATCH CONTINUATION" in calls[2]


def test_overflow_is_trimmed(patch_client):
    # Even if the model returns extra rows per batch, the total is trimmed to N.
    patch_client("over")
    out = _gen(30)
    assert len(out) == 30


def test_shortfall_is_padded(patch_client):
    # If the model keeps under-delivering, the exact count is still guaranteed
    # (padded by cloning) rather than returning a short set.
    out_client = patch_client("under")
    out = _gen(40)
    assert len(out) == 40
    # The loop is bounded — it does not spin forever on a stubborn model.
    assert len(out_client.chat.completions.calls) <= (
        ((40 + 19) // 20) * 2 + 3
    )


def test_small_request_single_batch(patch_client):
    # Small requests still work in one call (no regression for the common case).
    client = patch_client("exact")
    out = _gen(10)
    assert len(out) == 10
    assert len(client.chat.completions.calls) == 1
