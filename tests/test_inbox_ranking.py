import datetime as dt

from microsoft_mcp import tools as tools_mod
from microsoft_mcp.inbox_models import InboxItem
from microsoft_mcp.inbox_ranking import _compute_score, rank_items


def _future_iso(minutes: int) -> str:
    return (
        dt.datetime.now(dt.timezone.utc) + dt.timedelta(minutes=minutes)
    ).isoformat()


def test_inbox_item_creation():
    item = InboxItem(
        id="m1",
        kind="email",
        source_tool="list_emails",
        title="Test email",
    )
    assert item.id == "m1"
    assert item.kind == "email"
    assert item.score == 0.0


def test_rank_items_prioritizes_unread_over_read():
    items = [
        InboxItem(id="m1", kind="email", title="Read email", unread=False),
        InboxItem(id="m2", kind="email", title="Unread email", unread=True),
    ]
    ranked = rank_items(items)
    assert ranked[0].id == "m2"


def test_rank_items_prioritizes_soon_events():
    items = [
        InboxItem(id="m1", kind="email", title="FYI", unread=False),
        InboxItem(
            id="e1",
            kind="event",
            title="Starts soon",
            starts_in_minutes=10,
        ),
    ]
    ranked = rank_items(items)
    assert ranked[0].title == "Starts soon"


def test_rank_items_prioritizes_mentioned():
    items = [
        InboxItem(id="m1", kind="email", title="No mention", unread=True),
        InboxItem(
            id="m2", kind="email", title="Mentioned", unread=True, mentioned=True
        ),
    ]
    ranked = rank_items(items)
    assert ranked[0].id == "m2"


def test_rank_items_prioritizes_flagged():
    items = [
        InboxItem(id="m1", kind="email", title="Normal", unread=True),
        InboxItem(id="m2", kind="email", title="Flagged", unread=True, flagged=True),
    ]
    ranked = rank_items(items)
    assert ranked[0].id == "m2"


def test_rank_items_suppresses_newsletters():
    items = [
        InboxItem(id="m1", kind="email", title="Normal email", unread=True),
        InboxItem(
            id="m2", kind="email", title="Newsletter", unread=True, is_newsletter=True
        ),
    ]
    ranked = rank_items(items)
    # Newsletter should be ranked lower
    assert ranked[0].id == "m1"


def test_rank_items_returns_scores():
    items = [
        InboxItem(id="m1", kind="email", title="Test", unread=True),
    ]
    ranked = rank_items(items)
    assert ranked[0].score > 0


def test_inbox_item_to_dict():
    item = InboxItem(
        id="m1",
        kind="email",
        source_tool="list_emails",
        title="Test",
        snippet="preview text",
        web_url="https://example.com",
    )
    d = item.to_dict()
    assert d["id"] == "m1"
    assert d["kind"] == "email"
    assert d["title"] == "Test"
    assert "snippet" in d


def test_invite_message_populates_starts_in_minutes_under_15():
    raw = [
        {
            "id": "msg-1",
            "subject": "Imminent standup",
            "meetingMessageType": "meetingRequest",
            "startDateTime": {"dateTime": _future_iso(5)},
            "isRead": False,
        }
    ]
    items = tools_mod._invite_messages_to_inbox_items(raw)
    assert items[0].starts_in_minutes is not None
    assert items[0].starts_in_minutes <= 15
    # Ranker awards +25 (<=15 min meeting) on top of unread (+10) = >=35
    assert _compute_score(items[0]) >= 35


def test_event_populates_starts_in_minutes_1_to_2_hours():
    raw = [
        {
            "id": "evt-1",
            "subject": "Later meeting",
            "start": {"dateTime": _future_iso(90)},
        }
    ]
    items = tools_mod._events_to_inbox_items(raw)
    assert items[0].starts_in_minutes is not None
    assert 60 < items[0].starts_in_minutes <= 120
    # 60 < t <= 120 minutes -> proximity bucket 8.0
    assert _compute_score(items[0]) == 8.0


def test_event_within_7_days_still_scores_above_zero():
    """Regression: prior version capped proximity at 120 min and dropped to 0."""
    raw = [
        {
            "id": "evt-3d",
            "subject": "Mid-week sync",
            "start": {"dateTime": _future_iso(3 * 24 * 60)},  # 3 days out
        }
    ]
    items = tools_mod._events_to_inbox_items(raw)
    score = _compute_score(items[0])
    assert 0 < score < 3.0  # decays from 3.0 -> 0 across 1d..7d


def test_event_beyond_7_days_scores_zero():
    raw = [
        {
            "id": "evt-far",
            "subject": "Quarterly review",
            "start": {"dateTime": _future_iso(8 * 24 * 60)},
        }
    ]
    items = tools_mod._events_to_inbox_items(raw)
    assert _compute_score(items[0]) == 0.0


def test_past_events_have_none_starts_in_minutes():
    raw = [
        {
            "id": "evt-past",
            "subject": "Already happened",
            "start": {"dateTime": _future_iso(-30)},
        }
    ]
    items = tools_mod._events_to_inbox_items(raw)
    assert items[0].starts_in_minutes is None
    assert _compute_score(items[0]) == 0.0


def test_email_flagged_status_feeds_ranker():
    raw = [
        {
            "id": "m-1",
            "subject": "Action needed",
            "isRead": True,
            "flag": {"flagStatus": "flagged"},
        }
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert items[0].flagged is True
    assert _compute_score(items[0]) == 8.0


def test_email_not_flagged_when_status_missing_or_none():
    raw = [
        {
            "id": "m-2",
            "subject": "None",
            "isRead": True,
            "flag": {"flagStatus": "notFlagged"},
        },
        {"id": "m-3", "subject": "Missing", "isRead": True},
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert all(not i.flagged for i in items)


def test_newsletter_sender_heuristic_flags_item():
    raw = [
        {
            "id": "m-news",
            "subject": "Weekly digest",
            "isRead": False,
            "from": {
                "emailAddress": {"address": "noreply@substack.com", "name": "Substack"}
            },
        }
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert items[0].is_newsletter is True
    # unread(+10) + newsletter(-20) = -10
    assert _compute_score(items[0]) == -10.0


def test_human_sender_not_newsletter():
    raw = [
        {
            "id": "m-human",
            "subject": "Hey",
            "isRead": False,
            "from": {"emailAddress": {"address": "alice@company.com", "name": "Alice"}},
        }
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert items[0].is_newsletter is False


def test_mentioned_field_disabled_pending_replacement_signal():
    """`mentionsPreview` is not selectable on Graph v1.0 for our tenants.

    The previous behavior caused the whole list_inbox_items call to return
    zero emails (see mcp-tool-responses/v1/audit/inbox-triage). Field is
    now always False until a replacement mention signal is implemented.
    """
    raw = [
        {"id": "m-1", "subject": "a", "isRead": True},
        {
            "id": "m-2",
            "subject": "b",
            "isRead": True,
            "mentionsPreview": {"isMentioned": True},
        },
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert all(not i.mentioned for i in items)


def test_direct_to_boosts_score_above_cc_only(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "me@example.com")
    raw_direct = [
        {
            "id": "m-d",
            "subject": "Direct",
            "isRead": False,
            "toRecipients": [
                {"emailAddress": {"address": "me@example.com"}},
            ],
        }
    ]
    raw_cc = [
        {
            "id": "m-cc",
            "subject": "FYI cc",
            "isRead": False,
            "toRecipients": [
                {"emailAddress": {"address": "lead@example.com"}},
            ],
            "ccRecipients": [
                {"emailAddress": {"address": "me@example.com"}},
            ],
        }
    ]
    direct_item = tools_mod._emails_to_inbox_items(raw_direct)[0]
    cc_item = tools_mod._emails_to_inbox_items(raw_cc)[0]

    assert direct_item.direct_to is True
    assert direct_item.on_cc is False
    assert cc_item.direct_to is False
    assert cc_item.on_cc is True

    # direct (unread 10 + direct 5 = 15) > cc-only (unread 10 - cc 5 = 5)
    assert _compute_score(direct_item) == 15.0
    assert _compute_score(cc_item) == 5.0


_BOUNCE_TRUE_POSITIVES = [
    # subject prefixes seen in live triage + canonical MTA forms
    {
        "id": "b-1",
        "subject": "[Postmaster] Email Delivery Failure",
        "from": {"emailAddress": {"address": "postmaster@cresa.email"}},
    },
    {
        "id": "b-2",
        "subject": "[Postmaster] Messages on hold for broach@cresa.email",
        "from": {"emailAddress": {"address": "postmaster@cresa.email"}},
    },
    {
        "id": "b-3",
        "subject": "Undeliverable: oc office ti",
        "from": {"emailAddress": {"address": "cai.svcPOSTMASTER@coxautoinc.com"}},
    },
    {
        "id": "b-4",
        "subject": "Undeliverable: $1.08 otay mesa",
        "from": {"emailAddress": {"address": "postmaster@morfurniture.com"}},
    },
    {
        "id": "b-5",
        "subject": "Mail Delivery Failure",
        "from": {"emailAddress": {"address": "MAILER-DAEMON@example.org"}},
    },
    {
        "id": "b-6",
        "subject": "Delivery Status Notification (Failure)",
        "from": {"emailAddress": {"address": "mailer-daemon@googlemail.com"}},
    },
    {
        "id": "b-7",
        "subject": "Undelivered Mail Returned to Sender",
        "from": {"emailAddress": {"address": "postmaster@mailserver.local"}},
    },
    {
        "id": "b-8",
        "subject": "Failure Notice",
        "from": {"emailAddress": {"address": "MAILER-DAEMON@yahoo.com"}},
    },
]


_BOUNCE_TRUE_NEGATIVES = [
    # Real human/system mail that must NOT be flagged.
    {
        "name": "human conversation",
        "raw": {
            "id": "n-1",
            "subject": "Re: Q3 plan review",
            "from": {"emailAddress": {"address": "alice@acme.com"}},
        },
    },
    {
        "name": "newsletter with 'delivery' in subject",
        "raw": {
            "id": "n-2",
            "subject": "Same-day delivery now in your area!",
            "from": {"emailAddress": {"address": "marketing@retailer.com"}},
        },
    },
    {
        "name": "postmaster local-part but normal subject",
        "raw": {
            "id": "n-3",
            "subject": "Weekly DNS zone change request",
            "from": {"emailAddress": {"address": "postmaster@cresa.email"}},
        },
    },
    {
        "name": "DSN-style subject but human sender",
        "raw": {
            "id": "n-4",
            "subject": "Undeliverable: please reschedule",
            "from": {"emailAddress": {"address": "bob@vendor.com"}},
        },
    },
    {
        "name": "subject contains 'Failure Notice' deep inside, not at start",
        "raw": {
            "id": "n-5",
            "subject": "Action required after our Failure Notice last quarter",
            "from": {"emailAddress": {"address": "postmaster@cresa.email"}},
        },
    },
    {
        "name": "missing from field",
        "raw": {
            "id": "n-6",
            "subject": "Undeliverable: oc office ti",
        },
    },
    {
        "name": "empty subject with postmaster sender",
        "raw": {
            "id": "n-7",
            "subject": "",
            "from": {"emailAddress": {"address": "postmaster@cresa.email"}},
        },
    },
]


def test_is_bounce_detects_canonical_dsn_messages():
    """All known DSN/NDR shapes must be flagged."""
    items = tools_mod._emails_to_inbox_items(_BOUNCE_TRUE_POSITIVES)
    for raw, item in zip(_BOUNCE_TRUE_POSITIVES, items):
        assert item.is_bounce is True, (
            f"missed bounce: subject={raw['subject']!r} from={raw['from']!r}"
        )


def test_is_bounce_does_not_catch_legitimate_mail():
    """Strict AND logic — neither sender pattern alone nor subject pattern alone is enough."""
    raws = [case["raw"] for case in _BOUNCE_TRUE_NEGATIVES]
    items = tools_mod._emails_to_inbox_items(raws)
    for case, item in zip(_BOUNCE_TRUE_NEGATIVES, items):
        assert item.is_bounce is False, (
            f"false positive on case '{case['name']}': subject="
            f"{case['raw'].get('subject')!r} from="
            f"{case['raw'].get('from')!r}"
        )


def test_bounce_action_hints_are_delete_only():
    raw = _BOUNCE_TRUE_POSITIVES[0]
    item = tools_mod._emails_to_inbox_items([raw])[0]
    assert item.action_hints == ["delete"]


def test_bounce_score_sinks_below_human_mail(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "me@example.com")
    bounce_raw = {
        "id": "b-rank",
        "subject": "[Postmaster] Email Delivery Failure",
        "isRead": False,
        "from": {"emailAddress": {"address": "postmaster@cresa.email"}},
        "toRecipients": [{"emailAddress": {"address": "me@example.com"}}],
        "hasAttachments": True,
    }
    human_raw = {
        "id": "h-rank",
        "subject": "Re: contract",
        "isRead": False,
        "from": {"emailAddress": {"address": "alice@acme.com"}},
        "toRecipients": [{"emailAddress": {"address": "me@example.com"}}],
    }
    items = tools_mod._emails_to_inbox_items([bounce_raw, human_raw])
    ranked = rank_items(items)
    # Human mail must come first; bounce sinks below zero.
    assert ranked[0].id == "h-rank"
    assert ranked[1].id == "b-rank"
    assert ranked[1].score < 0
    assert "bounce" in ranked[1].reason


def test_bcc_only_penalised_more_than_cc(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "me@example.com")
    raw_bcc = [
        {
            "id": "m-bcc",
            "subject": "Quiet copy",
            "isRead": False,
            "toRecipients": [
                {"emailAddress": {"address": "lead@example.com"}},
            ],
            "bccRecipients": [
                {"emailAddress": {"address": "me@example.com"}},
            ],
        }
    ]
    item = tools_mod._emails_to_inbox_items(raw_bcc)[0]
    assert item.on_bcc is True
    # unread 10 - bcc 8 = 2
    assert _compute_score(item) == 2.0


def test_action_hints_include_reply_when_direct_unread(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "me@example.com")
    raw = [
        {
            "id": "m-r",
            "subject": "Need answer",
            "isRead": False,
            "toRecipients": [
                {"emailAddress": {"address": "me@example.com"}},
            ],
        }
    ]
    item = tools_mod._emails_to_inbox_items(raw)[0]
    assert "reply" in item.action_hints


def test_action_hints_newsletter_says_archive_unsubscribe():
    raw = [
        {
            "id": "m-news",
            "subject": "Weekly digest",
            "isRead": False,
            "from": {
                "emailAddress": {"address": "noreply@substack.com", "name": "Substack"}
            },
        }
    ]
    item = tools_mod._emails_to_inbox_items(raw)[0]
    assert item.action_hints == ["archive", "unsubscribe"]


def test_reason_string_populated_after_ranking():
    item = InboxItem(
        id="m1",
        kind="email",
        title="x",
        unread=True,
        flagged=True,
    )
    ranked = rank_items([item])
    assert "unread" in ranked[0].reason
    assert "flagged" in ranked[0].reason


def test_to_dict_aliases_sender_from_first_participant():
    item = InboxItem(
        id="m1",
        kind="email",
        title="x",
        participants=["Alice <a@example.com>"],
    )
    d = item.to_dict()
    assert d["sender"] == "Alice <a@example.com>"


def test_to_dict_emits_cc_and_bcc_when_present():
    item = InboxItem(
        id="m1",
        kind="email",
        title="x",
        cc=["Bob <b@example.com>"],
        bcc=["Carol <c@example.com>"],
    )
    d = item.to_dict()
    assert d["cc"] == ["Bob <b@example.com>"]
    assert d["bcc"] == ["Carol <c@example.com>"]
