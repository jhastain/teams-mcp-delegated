# SPDX-FileCopyrightText: 2026 Jason Hastain
# SPDX-License-Identifier: Apache-2.0
"""Unit tests for attachment extraction (adaptive cards and O365 connector cards)."""
import json
from types import SimpleNamespace

from mcp_teams_server.teams import (
    _extract_adaptive_card_text,
    _extract_attachments,
    _extract_o365_connector_card_text,
)


def _make_attachment(content_type: str, content: str) -> SimpleNamespace:
    return SimpleNamespace(content_type=content_type, content=content)


# ---------------------------------------------------------------------------
# O365 Connector Card extraction
# ---------------------------------------------------------------------------

AZURE_DEVOPS_CARD = {
    "summary": "Build Run Checkout-API - failed",
    "themeColor": "#da0a00",
    "sections": [
        {
            "text": '<p>Build for <a href="https://dev.azure.com/org/proj/_git/Repo">Repo</a>, finished 03/25/2026</p>',
            "activityTitle": '<p><a href="https://dev.azure.com/org/proj/_build/results?buildId=123">[Pipeline] Run</a> (Failed)</p>',
            "activitySubtitle": "<p>Trigger: manual</p>",
            "markdown": True,
        }
    ],
    "potentialAction": [],
}

SENTRY_CARD = {
    "summary": "ISSUES FOUND - Sentry Release 4821-rc (web-api-dotnet)",
    "themeColor": "FF6B6B",
    "sections": [
        {
            "facts": [
                {"name": "Release Version", "value": "<p>4821-rc</p>"},
                {"name": "Issues Found", "value": "<p>3</p>"},
                {"name": "Project", "value": "<p>web-api-dotnet</p>"},
            ],
            "activityTitle": "<p>ISSUES FOUND - Deployment Candidate</p>",
            "activitySubtitle": "<p>issues detected</p>",
        },
        {
            "text": "<p><strong>Issues Found:</strong></p>",
        },
    ],
    "potentialAction": [
        {
            "targets": [{"os": "default", "uri": "https://sentry.io/issues/"}],
            "@type": "OpenUri",
            "name": "View Issues in Sentry",
        }
    ],
}


def test_o365_azure_devops_card():
    text = _extract_o365_connector_card_text(AZURE_DEVOPS_CARD)
    assert "Build Run Checkout-API - failed" in text
    assert "[Pipeline] Run (Failed)" in text
    assert "Trigger: manual" in text
    assert "finished 03/25/2026" in text
    # HTML tags stripped
    assert "<p>" not in text
    assert "<a " not in text


def test_o365_sentry_card():
    text = _extract_o365_connector_card_text(SENTRY_CARD)
    assert "ISSUES FOUND" in text
    assert "Release Version: 4821-rc" in text
    assert "Issues Found: 3" in text
    assert "Project: web-api-dotnet" in text
    assert "[View Issues in Sentry](https://sentry.io/issues/)" in text
    assert "<p>" not in text


def test_o365_empty_card():
    assert _extract_o365_connector_card_text({}) == ""


def test_o365_summary_only():
    text = _extract_o365_connector_card_text({"summary": "Alert fired"})
    assert text == "Alert fired"


def test_o365_facts_without_name():
    card = {
        "sections": [{"facts": [{"name": "", "value": "solo value"}]}]
    }
    text = _extract_o365_connector_card_text(card)
    assert "solo value" in text


def test_o365_action_without_uri():
    card = {"potentialAction": [{"name": "Click me"}]}
    text = _extract_o365_connector_card_text(card)
    assert "[Click me]" in text


# ---------------------------------------------------------------------------
# _extract_attachments dispatches to correct extractor
# ---------------------------------------------------------------------------

def test_extract_attachments_adaptive_card():
    card = {"body": [{"type": "TextBlock", "text": "Hello from adaptive"}]}
    att = _make_attachment(
        "application/vnd.microsoft.card.adaptive", json.dumps(card)
    )
    results = _extract_attachments([att])
    assert len(results) == 1
    assert results[0].type == "adaptive_card"
    assert "Hello from adaptive" in results[0].extracted_text


def test_extract_attachments_o365_connector():
    att = _make_attachment(
        "application/vnd.microsoft.teams.card.o365connector",
        json.dumps(AZURE_DEVOPS_CARD),
    )
    results = _extract_attachments([att])
    assert len(results) == 1
    assert results[0].type == "o365_connector_card"
    assert "Build Run Checkout-API - failed" in results[0].extracted_text


def test_extract_attachments_mixed():
    adaptive = {"body": [{"type": "TextBlock", "text": "adaptive text"}]}
    atts = [
        _make_attachment(
            "application/vnd.microsoft.card.adaptive", json.dumps(adaptive)
        ),
        _make_attachment(
            "application/vnd.microsoft.teams.card.o365connector",
            json.dumps(SENTRY_CARD),
        ),
        _make_attachment("reference", "not-json"),  # non-card attachment
    ]
    results = _extract_attachments(atts)
    assert len(results) == 2
    types = {r.type for r in results}
    assert types == {"adaptive_card", "o365_connector_card"}


def test_extract_attachments_empty():
    assert _extract_attachments(None) == []
    assert _extract_attachments([]) == []


def test_extract_attachments_invalid_json():
    att = _make_attachment(
        "application/vnd.microsoft.teams.card.o365connector", "not json"
    )
    assert _extract_attachments([att]) == []


# ---------------------------------------------------------------------------
# Adaptive card extraction (regression tests from v1.0.8)
# ---------------------------------------------------------------------------

def test_adaptive_textblock():
    body = [{"type": "TextBlock", "text": "Line 1"}, {"type": "TextBlock", "text": "Line 2"}]
    assert _extract_adaptive_card_text(body) == "Line 1\nLine 2"


def test_adaptive_factset():
    body = [{"type": "FactSet", "facts": [{"title": "Key", "value": "Val"}]}]
    assert "Key: Val" in _extract_adaptive_card_text(body)


def test_adaptive_nested_container():
    body = [{"type": "Container", "items": [{"type": "TextBlock", "text": "nested"}]}]
    assert "nested" in _extract_adaptive_card_text(body)


def test_adaptive_empty_body():
    assert _extract_adaptive_card_text([]) == ""
