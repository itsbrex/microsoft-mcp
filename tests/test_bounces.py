"""Tests for bounces.py — NDR classifier, DSN parser, folder scanner, and CSV export.

Import convention follows the repo pattern:
    from src.microsoft_mcp import <module>
"""

import csv
import io
from unittest import mock

import pytest

from src.microsoft_mcp import bounces
from src.microsoft_mcp.bounces import (
    classify_bounce_message,
    determine_bounce_reason,
    extract_email_from_text,
    is_bounce_message,
    iter_folder_messages,
    parse_dsn_content,
    parse_name_from_email,
    scan_folder,
    write_csv,
)

# ---------------------------------------------------------------------------
# Sample DSN block
# ---------------------------------------------------------------------------

SAMPLE_DSN = """\
Reporting-MTA: dns; mail.example.com

Final-Recipient: rfc822; user@example.com
Action: failed
Status: 5.1.1
Diagnostic-Code: smtp; 550 5.1.1 The email account that you tried to reach does not exist.
"""

# ---------------------------------------------------------------------------
# extract_email_from_text
# ---------------------------------------------------------------------------


class TestExtractEmailFromText:
    def test_extracts_plain_address(self):
        result = extract_email_from_text(
            "Please contact jane.doe@example.com for help."
        )
        assert result == "jane.doe@example.com"

    def test_returns_none_for_empty(self):
        assert extract_email_from_text("") is None

    def test_returns_none_for_none(self):
        assert extract_email_from_text(None) is None

    def test_returns_none_for_no_address(self):
        assert extract_email_from_text("no email address here") is None

    def test_extracts_first_of_multiple(self):
        result = extract_email_from_text("From: a@x.com, To: b@y.com")
        assert result == "a@x.com"


# ---------------------------------------------------------------------------
# parse_name_from_email
# ---------------------------------------------------------------------------


class TestParseNameFromEmail:
    def test_dot_separated(self):
        first, last = parse_name_from_email("jane.doe@x.com")
        assert first == "Jane"
        assert last == "Doe"

    def test_underscore_separated(self):
        first, last = parse_name_from_email("john_smith@corp.com")
        assert first == "John"
        assert last == "Smith"

    def test_hyphen_separated(self):
        first, last = parse_name_from_email("mary-jones@example.org")
        assert first == "Mary"
        assert last == "Jones"

    def test_none_input(self):
        assert parse_name_from_email(None) == ("", "")

    def test_no_at_sign(self):
        assert parse_name_from_email("notanemail") == ("", "")

    def test_single_part_local(self):
        first, last = parse_name_from_email("alice@example.com")
        assert first == "Alice"
        assert last == ""


# ---------------------------------------------------------------------------
# is_bounce_message
# ---------------------------------------------------------------------------


class TestIsBounceMessage:
    def test_postmaster_undeliverable_is_bounce(self):
        result = is_bounce_message(
            subject="Undeliverable: your message to bob@corp.com",
            sender_email="postmaster@mailserver.com",
        )
        assert result is True

    def test_mailer_daemon_delivery_failure_is_bounce(self):
        result = is_bounce_message(
            subject="Mail Delivery Failed",
            sender_email="mailer-daemon@example.com",
        )
        assert result is True

    def test_normal_email_is_not_bounce(self):
        result = is_bounce_message(
            subject="Meeting tomorrow at 3pm",
            sender_email="colleague@company.com",
        )
        assert result is False

    def test_strong_subject_indicator_alone_triggers(self):
        # Even without a postmaster sender, strong subject => bounce
        result = is_bounce_message(
            subject="Undeliverable: your message",
            sender_email="someone@company.com",
        )
        assert result is True

    def test_excluded_prefix_auto_reply_not_bounce(self):
        result = is_bounce_message(
            subject="Automatic reply: I am out of the office",
            sender_email="postmaster@mail.com",
        )
        assert result is False

    def test_body_pattern_triggers_when_use_body_true(self):
        result = is_bounce_message(
            subject="Message returned",
            sender_email="noreply@mail.com",
            body="Delivery has failed to these recipients: bob@corp.com",
            use_body=True,
        )
        assert result is True

    def test_body_pattern_ignored_when_use_body_false(self):
        result = is_bounce_message(
            subject="Hello",
            sender_email="noreply@mail.com",
            body="Delivery has failed to these recipients: bob@corp.com",
            use_body=False,
        )
        assert result is False

    def test_delivery_status_notification_subject(self):
        result = is_bounce_message(
            subject="Delivery Status Notification (Failure)",
            sender_email="system@mail.com",
        )
        assert result is True


# ---------------------------------------------------------------------------
# determine_bounce_reason
# ---------------------------------------------------------------------------


class TestDetermineBounceReason:
    def test_550_511_maps_to_invalid_recipient(self):
        reason = determine_bounce_reason(
            subject="Undeliverable: test",
            body="550 5.1.1 The email account that you tried to reach does not exist.",
        )
        assert reason == "Invalid Recipient"

    def test_mailbox_full_reason(self):
        reason = determine_bounce_reason(
            subject="Mail Delivery Failed",
            body="552 5.2.2 Mailbox Full — storage quota exceeded.",
        )
        assert reason == "Mailbox Full (552 5.2.2)"

    def test_relay_access_denied(self):
        reason = determine_bounce_reason(
            subject="Failed",
            body="Relay access denied. Cannot route to that domain.",
        )
        assert reason == "Relay Access Denied"

    def test_unknown_reason_fallback(self):
        reason = determine_bounce_reason(
            subject="Something bounced",
            body="No recognizable error code here.",
        )
        assert reason == "Unknown"

    def test_undeliverable_subject_fallback(self):
        reason = determine_bounce_reason(
            subject="Undeliverable message",
            body="No specific SMTP code present.",
        )
        assert reason == "Undeliverable"

    def test_failure_subject_fallback(self):
        reason = determine_bounce_reason(
            subject="Delivery Failure today",
            body="No specific code.",
        )
        assert reason == "Delivery Failure (Unspecified)"

    def test_code_in_subject_is_matched(self):
        # 550 5.1.1 in subject (not body) should still match
        reason = determine_bounce_reason(
            subject="550 5.1.1 user not found",
            body="",
        )
        assert reason == "Invalid Recipient"


# ---------------------------------------------------------------------------
# parse_dsn_content
# ---------------------------------------------------------------------------


class TestParseDsnContent:
    def test_extracts_final_recipient(self):
        result = parse_dsn_content(SAMPLE_DSN)
        assert result["final_recipient"] == "user@example.com"

    def test_extracts_status(self):
        result = parse_dsn_content(SAMPLE_DSN)
        assert result["status"] == "5.1.1"

    def test_extracts_action(self):
        result = parse_dsn_content(SAMPLE_DSN)
        assert result["action"] == "failed"

    def test_extracts_diagnostic_code(self):
        result = parse_dsn_content(SAMPLE_DSN)
        assert result["diagnostic_code"] is not None
        assert "550" in result["diagnostic_code"]

    def test_returns_all_required_keys(self):
        result = parse_dsn_content(SAMPLE_DSN)
        assert set(result.keys()) == {
            "final_recipient",
            "action",
            "status",
            "diagnostic_code",
            "display_name",
        }

    def test_display_name_none_when_absent(self):
        result = parse_dsn_content(SAMPLE_DSN)
        assert result["display_name"] is None

    def test_x_failed_recipients_takes_priority(self):
        content = "X-Failed-Recipients: priority@example.com\nFinal-Recipient: rfc822; secondary@example.com\n"
        result = parse_dsn_content(content)
        assert result["final_recipient"] == "priority@example.com"

    def test_empty_string(self):
        result = parse_dsn_content("")
        assert result["final_recipient"] is None
        assert result["status"] is None

    def test_original_recipient_fallback(self):
        content = "Original-Recipient: rfc822; fallback@example.com\nStatus: 5.1.1\n"
        result = parse_dsn_content(content)
        assert result["final_recipient"] == "fallback@example.com"

    def test_display_name_extracted(self):
        content = "X-Display-Name: Bob Smith\nStatus: 5.1.1\n"
        result = parse_dsn_content(content)
        assert result["display_name"] == "Bob Smith"


# ---------------------------------------------------------------------------
# classify_bounce_message
# ---------------------------------------------------------------------------

BOUNCE_MSG_DICT: dict = {
    "id": "AABounce123",
    "subject": "Undeliverable: Hello bob",
    "from": {
        "emailAddress": {
            "address": "postmaster@mailserver.com",
            "name": "Postmaster",
        }
    },
    "receivedDateTime": "2025-03-15T10:30:00Z",
    "hasAttachments": False,
    "body": {
        "contentType": "text",
        "content": (
            "Your message to bob.jones@company.com could not be delivered.\n"
            "550 5.1.1 The email account does not exist.\n"
        ),
    },
}

NORMAL_MSG_DICT: dict = {
    "id": "AANormal456",
    "subject": "Lunch plans",
    "from": {
        "emailAddress": {
            "address": "alice@company.com",
            "name": "Alice",
        }
    },
    "receivedDateTime": "2025-03-15T12:00:00Z",
    "hasAttachments": False,
    "body": {
        "contentType": "text",
        "content": "Let's grab lunch at noon today!",
    },
}


class TestClassifyBounceMessage:
    def test_returns_none_for_normal_message(self):
        result = classify_bounce_message(NORMAL_MSG_DICT)
        assert result is None

    def test_returns_dict_for_bounce(self):
        result = classify_bounce_message(BOUNCE_MSG_DICT)
        assert result is not None
        assert isinstance(result, dict)

    def test_all_required_keys_present(self):
        result = classify_bounce_message(BOUNCE_MSG_DICT)
        assert result is not None
        expected_keys = {
            "first_name",
            "last_name",
            "email",
            "reason",
            "date",
            "iso_date",
            "subject",
            "sender",
            "body",
            "message_id",
            "has_attachments",
        }
        assert expected_keys == set(result.keys())

    def test_email_extracted_from_body(self):
        result = classify_bounce_message(BOUNCE_MSG_DICT)
        assert result is not None
        assert result["email"] == "bob.jones@company.com"

    def test_name_parsed_from_email(self):
        result = classify_bounce_message(BOUNCE_MSG_DICT)
        assert result is not None
        assert result["first_name"] == "Bob"
        assert result["last_name"] == "Jones"

    def test_reason_derived_from_smtp_code(self):
        result = classify_bounce_message(BOUNCE_MSG_DICT)
        assert result is not None
        assert result["reason"] == "Invalid Recipient"

    def test_iso_date_from_received_date_time(self):
        result = classify_bounce_message(BOUNCE_MSG_DICT)
        assert result is not None
        assert result["iso_date"] == "2025-03-15T10:30:00Z"
        assert result["date"] == "2025-03-15T10:30:00Z"

    def test_message_id_preserved(self):
        result = classify_bounce_message(BOUNCE_MSG_DICT)
        assert result is not None
        assert result["message_id"] == "AABounce123"

    def test_has_attachments_false(self):
        result = classify_bounce_message(BOUNCE_MSG_DICT)
        assert result is not None
        assert result["has_attachments"] is False

    def test_subject_preserved(self):
        result = classify_bounce_message(BOUNCE_MSG_DICT)
        assert result is not None
        assert result["subject"] == "Undeliverable: Hello bob"

    def test_sender_preserved(self):
        result = classify_bounce_message(BOUNCE_MSG_DICT)
        assert result is not None
        assert result["sender"] == "postmaster@mailserver.com"

    def test_no_recipient_in_message_uses_fallback(self):
        msg = {
            "id": "AANoRecip",
            "subject": "Delivery Failure",
            "from": {
                "emailAddress": {
                    "address": "mailer-daemon@mail.example.com",
                    "name": "Mail Daemon",
                }
            },
            "receivedDateTime": "2025-04-01T09:00:00Z",
            "hasAttachments": False,
            "body": {
                "contentType": "text",
                "content": "Mail Delivery Failed. No address found.",
            },
        }
        result = classify_bounce_message(msg)
        assert result is not None
        assert result["email"] == "(not extracted)"
        assert result["first_name"] == "Unknown"
        assert result["last_name"] == "Recipient"


# ---------------------------------------------------------------------------
# Module-level constants sanity checks
# ---------------------------------------------------------------------------


class TestModuleConstants:
    def test_subject_keywords_is_list_of_strings(self):
        assert isinstance(bounces.SUBJECT_KEYWORDS, list)
        assert all(isinstance(k, str) for k in bounces.SUBJECT_KEYWORDS)

    def test_sender_patterns_is_list_of_strings(self):
        assert isinstance(bounces.SENDER_PATTERNS, list)
        assert all(isinstance(p, str) for p in bounces.SENDER_PATTERNS)

    def test_body_patterns_is_list_of_strings(self):
        assert isinstance(bounces.BODY_PATTERNS, list)
        assert all(isinstance(p, str) for p in bounces.BODY_PATTERNS)

    def test_bounce_reasons_is_list_of_tuples(self):
        assert isinstance(bounces.BOUNCE_REASONS, list)
        for item in bounces.BOUNCE_REASONS:
            assert isinstance(item, tuple)
            assert len(item) == 2

    def test_strong_subject_indicators_is_tuple(self):
        assert isinstance(bounces.STRONG_SUBJECT_INDICATORS, tuple)

    def test_excluded_subject_prefixes_is_tuple(self):
        assert isinstance(bounces.EXCLUDED_SUBJECT_PREFIXES, tuple)


# ---------------------------------------------------------------------------
# Parametrized regression — every BOUNCE_REASONS entry must be reachable
# ---------------------------------------------------------------------------
#
# Matching is FIRST-MATCH-WINS (top-to-bottom).  Each trigger string below is
# crafted to hit exactly the intended entry without being swallowed by an
# earlier one.  Ordering notes that matter:
#
#   • Entry 1 (550 5.1.10) vs Entry 2 (550 5.1.1\b): the \b after the second
#     "1" requires a non-word char at that position.  "550 5.1.10" has digit
#     "0" after the second "1", so \b does NOT match — Entry 1 is safely
#     distinct from Entry 2.
#   • Entry 6 (552 5.2.2 → "Mailbox Full (552 5.2.2)") vs Entry 26 ("Mailbox
#     full" → "Mailbox Full"): trigger Entry 26 without "552 5.2.2".
#   • Entry 5 (550 5.2.1 → "Mailbox Disabled (550 5.2.1)") vs Entry 27
#     ("Mailbox disabled" → "Mailbox Disabled"): trigger Entry 27 without
#     "550 5.2.1".
#   • Entry 38 ("Delivery not authorized" → "Delivery Not Authorized") is
#     reachable because Entry 4 ("550 5.7.1") matches on the SMTP code, not
#     on the phrase.  A body with only the phrase (no code) reaches Entry 38.


@pytest.mark.parametrize(
    "text, expected_reason",
    [
        # --- SMTP status codes (entries 1–6) ---
        # Entry 1: 550 5.1.10  (must come before 5.1.1\b — digit "0" blocks \b)
        ("550 5.1.10 Recipient not found", "Recipient Not Found (550 5.1.10)"),
        # Entry 2: 550 5.1.1\b  (no trailing digit, so \b fires)
        ("550 5.1.1 The account does not exist", "Invalid Recipient"),
        # Entry 3: 550 5.4.1
        ("550 5.4.1 No answer from host", "No Answer from Host (550 5.4.1)"),
        # Entry 4: 550 5.7.1
        (
            "550 5.7.1 Delivery not authorized by policy",
            "Delivery Not Authorized (550 5.7.1)",
        ),
        # Entry 5: 550 5.2.1
        ("550 5.2.1 Mailbox disabled for this user", "Mailbox Disabled (550 5.2.1)"),
        # Entry 6: 552 5.2.2
        ("552 5.2.2 Over quota", "Mailbox Full (552 5.2.2)"),
        # --- Mail loop (entries 7–8) ---
        # Entry 7: 5.4.14
        (
            "Diagnostic-Code: smtp; 5.4.14 Hop count exceeded — mail loop",
            "Mail Loop Detected (5.4.14)",
        ),
        # Entry 8: hop count exceeded (text only, no "5.4.14")
        ("Hop count exceeded; message looped too many times", "Mail Loop Detected"),
        # --- Exchange resolver codes (entries 9–10) ---
        # Entry 9: RESOLVER.ADR.RecipientNotFound
        (
            "RESOLVER.ADR.RecipientNotFound; recipient lookup failed",
            "Recipient Not Found",
        ),
        # Entry 10: RESOLVER.ADR.BadPrimary
        ("RESOLVER.ADR.BadPrimary; primary SMTP address is bad", "Bad Primary Address"),
        # --- Connection / network (entries 11–14) ---
        # Entry 11: Communications error
        ("Communications error occurred during SMTP handshake", "Communications Error"),
        # Entry 12: Read timed out
        ("Read timed out waiting for banner", "Read Timeout"),
        # Entry 13: Connection timed out  (must not contain "Read timed out")
        ("Connection timed out while connecting to MX", "Connection Timeout"),
        # Entry 14: Connection refused
        ("Connection refused by remote host on port 25", "Connection Refused"),
        # Entry 15: Network unreachable
        ("Network unreachable — no route to 203.0.113.1", "Network Unreachable"),
        # --- DNS / MX (entries 16–19) ---
        # Entry 16: MX records? (?:or )?is invalid
        ("The MX records is invalid for this domain", "Invalid MX Records"),
        # Entry 17: Domain has no MX records
        ("Domain has no MX records configured", "No MX Records"),
        # Entry 18: DNS lookup failed
        ("DNS lookup failed for recipient domain", "DNS Lookup Failed"),
        # Entry 19: Temporary error looking up MX
        ("Temporary error looking up MX for example.com", "Temporary MX Lookup Error"),
        # --- Recipient (entries 20–24) ---
        # Entry 20: Recipient email address is possibly incorrect
        ("Recipient email address is possibly incorrect", "Invalid Recipient Address"),
        # Entry 21: User unknown
        ("User unknown in local recipient table", "User Unknown"),
        # Entry 22: No such user
        ("No such user here", "No Such User"),
        # Entry 23: Address not found
        ("Address not found in directory", "Address Not Found"),
        # Entry 24: not found by SMTP address lookup
        ("550 not found by SMTP address lookup", "SMTP Address Not Found"),
        # --- Mailbox (entries 25–29) ---
        # Entry 25: Mailbox unavailable
        ("Mailbox unavailable or access denied", "Mailbox Unavailable"),
        # Entry 26: Mailbox full  (no "552 5.2.2" to avoid Entry 6)
        ("Mailbox full — the recipient's storage is at capacity", "Mailbox Full"),
        # Entry 27: Mailbox disabled  (no "550 5.2.1" to avoid Entry 5)
        ("Mailbox disabled by administrator policy", "Mailbox Disabled"),
        # Entry 28: Quota exceeded
        ("Quota exceeded for this mailbox", "Quota Exceeded"),
        # Entry 29: Insufficient storage
        ("Insufficient storage on server", "Insufficient Storage"),
        # --- Size / content (entry 30) ---
        # Entry 30: Message too large
        ("Message too large for server to accept", "Message Too Large"),
        # --- Relay / policy (entries 31–38) ---
        # Entry 31: Relay access denied
        ("Relay access denied from this IP", "Relay Access Denied"),
        # Entry 32: Sender address rejected
        ("Sender address rejected: domain not found", "Sender Rejected"),
        # Entry 33: Recipient address rejected
        ("Recipient address rejected: access denied", "Recipient Rejected"),
        # Entry 34: Blocked by recipient
        ("Blocked by recipient's mail filter", "Blocked by Recipient"),
        # Entry 35: Spam filter
        ("Spam filter triggered by message content", "Spam Filter"),
        # Entry 36: Content filter
        ("Content filter rejected the message", "Content Filter"),
        # Entry 37: Policy rejection
        ("Policy rejection — message violates domain policy", "Policy Rejection"),
        # Entry 38: Delivery not authorized  (plain phrase, no "550 5.7.1")
        (
            "Delivery not authorized by the destination server",
            "Delivery Not Authorized",
        ),
        # --- Fallback / unknown ---
        ("This text matches no known bounce pattern at all", "Unknown"),
    ],
)
def test_determine_bounce_reason_all_patterns(text: str, expected_reason: str) -> None:
    """Every BOUNCE_REASONS entry must be reachable via a crafted trigger string.

    The trigger is passed as the *body* argument so the subject remains neutral
    and cannot accidentally fire an earlier pattern.
    """
    assert bounces.determine_bounce_reason("", text) == expected_reason


# ---------------------------------------------------------------------------
# Sample Graph message fixtures
# ---------------------------------------------------------------------------

_BOUNCE_MSG = {
    "id": "AABounce001",
    "subject": "Undeliverable: Hello bob",
    "from": {
        "emailAddress": {
            "address": "postmaster@mailserver.com",
            "name": "Postmaster",
        }
    },
    "receivedDateTime": "2025-03-15T10:30:00Z",
    "hasAttachments": False,
    "body": {
        "contentType": "text",
        "content": (
            "Your message to bob.jones@company.com could not be delivered.\n"
            "550 5.1.1 The email account does not exist.\n"
        ),
    },
}

_NORMAL_MSG = {
    "id": "AANormal001",
    "subject": "Lunch plans",
    "from": {"emailAddress": {"address": "alice@company.com", "name": "Alice"}},
    "receivedDateTime": "2025-03-15T12:00:00Z",
    "hasAttachments": False,
    "body": {"contentType": "text", "content": "Let's grab lunch at noon!"},
}


# ---------------------------------------------------------------------------
# iter_folder_messages
# ---------------------------------------------------------------------------


class TestIterFolderMessages:
    """Tests for pagination and limit behaviour of iter_folder_messages."""

    def _make_request(self, pages: list[dict]) -> object:
        """Return a fake request callable that serves ``pages`` in sequence."""
        calls: list[tuple] = []
        page_iter = iter(pages)

        def fake_request(method, path, **kwargs):
            calls.append((method, path, kwargs))
            try:
                return next(page_iter)
            except StopIteration:
                return {"value": []}

        fake_request.calls = calls  # type: ignore[attr-defined]
        return fake_request

    def test_single_page_no_next_link(self):
        page = {"value": [_BOUNCE_MSG, _NORMAL_MSG]}
        req = self._make_request([page])
        result = list(iter_folder_messages(req, "inbox"))
        assert len(result) == 2

    def test_two_pages_via_next_link(self):
        """Page 1 includes @odata.nextLink; page 2 does not. Both pages yielded."""
        page1 = {
            "value": [_BOUNCE_MSG],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skip=50",
        }
        page2 = {"value": [_NORMAL_MSG]}
        calls: list[tuple] = []

        def fake_request(method, path, **kwargs):
            calls.append((method, path))
            if len(calls) == 1:
                return page1
            return page2

        result = list(iter_folder_messages(fake_request, "inbox"))
        assert len(result) == 2
        # Second call must use the base-stripped path, not the full URL
        assert calls[1][1] == "/me/mailFolders/inbox/messages?$skip=50"

    def test_next_link_graph_base_stripped(self):
        """_strip_graph_base is applied: the second call must not start with https://."""
        next_link = (
            "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skip=50"
        )
        page1 = {"value": [_BOUNCE_MSG], "@odata.nextLink": next_link}
        page2 = {"value": []}
        second_path: list[str] = []

        def fake_request(method, path, **kwargs):
            if len(second_path) == 0 and path.startswith(
                "/me/mailFolders/inbox/messages?"
            ):
                second_path.append(path)
            return page1 if not second_path else page2

        list(iter_folder_messages(fake_request, "inbox"))
        assert second_path, "second call was never made"
        assert not second_path[0].startswith("https://")
        assert second_path[0] == "/me/mailFolders/inbox/messages?$skip=50"

    def test_limit_caps_yielded_messages(self):
        """limit=1 must stop after yielding one message even if more exist."""
        page = {"value": [_BOUNCE_MSG, _NORMAL_MSG, _BOUNCE_MSG]}
        req = self._make_request([page])
        result = list(iter_folder_messages(req, "inbox", limit=1))
        assert len(result) == 1

    def test_limit_across_pages(self):
        """limit=1 must stop after page 1 even when page 1 has a nextLink."""
        page1 = {
            "value": [_BOUNCE_MSG],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skip=50",
        }
        page2 = {"value": [_NORMAL_MSG]}
        call_count = [0]

        def fake_request(method, path, **kwargs):
            call_count[0] += 1
            return page1 if call_count[0] == 1 else page2

        result = list(iter_folder_messages(fake_request, "inbox", limit=1))
        assert len(result) == 1
        assert call_count[0] == 1  # page 2 was never fetched


# ---------------------------------------------------------------------------
# scan_folder
# ---------------------------------------------------------------------------


class TestScanFolder:
    """Tests for scan_folder — filters bounces from a mix of messages."""

    def test_returns_only_bounces(self):
        """A mix of bounce + normal messages: only bounces are classified."""
        page = {"value": [_BOUNCE_MSG, _NORMAL_MSG]}

        def fake_request(method, path, **kwargs):
            return page

        rows = scan_folder(fake_request, "inbox")
        assert len(rows) == 1
        assert rows[0]["message_id"] == "AABounce001"

    def test_no_bounces_returns_empty(self):
        page = {"value": [_NORMAL_MSG]}

        def fake_request(method, path, **kwargs):
            return page

        rows = scan_folder(fake_request, "inbox")
        assert rows == []

    def test_limit_passed_to_iter(self):
        """scan_folder's limit caps messages SCANNED, not just returned."""
        page = {
            "value": [_BOUNCE_MSG, _BOUNCE_MSG, _BOUNCE_MSG],
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$skip=50",
        }
        call_count = [0]

        def fake_request(method, path, **kwargs):
            call_count[0] += 1
            return page if call_count[0] == 1 else {"value": []}

        rows = scan_folder(fake_request, "inbox", limit=2)
        assert len(rows) == 2
        # Should NOT have fetched page 2 (limit=2 hit on page 1)
        assert call_count[0] == 1

    def test_classified_record_has_expected_keys(self):
        page = {"value": [_BOUNCE_MSG]}

        def fake_request(method, path, **kwargs):
            return page

        rows = scan_folder(fake_request, "inbox")
        assert rows
        expected_keys = {
            "first_name",
            "last_name",
            "email",
            "reason",
            "date",
            "iso_date",
            "subject",
            "sender",
            "body",
            "message_id",
            "has_attachments",
        }
        assert set(rows[0].keys()) == expected_keys


# ---------------------------------------------------------------------------
# write_csv
# ---------------------------------------------------------------------------


class TestWriteCsv:
    def test_writes_header_and_row(self, tmp_path):
        row = {
            "first_name": "Bob",
            "last_name": "Jones",
            "email": "bob.jones@company.com",
            "reason": "Invalid Recipient",
            "date": "2025-03-15T10:30:00Z",
            "iso_date": "2025-03-15T10:30:00Z",
            "subject": "Undeliverable: Hello bob",
            "sender": "postmaster@mailserver.com",
            "body": "Your message could not be delivered.",
            "message_id": "AABounce001",
            "has_attachments": False,
        }
        out = tmp_path / "out.csv"
        write_csv([row], out)
        content = out.read_text(encoding="utf-8")
        reader = csv.DictReader(io.StringIO(content))
        rows = list(reader)
        assert len(rows) == 1
        assert rows[0]["email"] == "bob.jones@company.com"
        assert rows[0]["reason"] == "Invalid Recipient"

    def test_empty_rows_writes_header_only(self, tmp_path):
        out = tmp_path / "empty.csv"
        write_csv([], out)
        content = out.read_text(encoding="utf-8")
        reader = csv.DictReader(io.StringIO(content))
        rows = list(reader)
        assert rows == []
        # Header must still be present
        assert reader.fieldnames is not None
        assert "email" in reader.fieldnames
        assert "reason" in reader.fieldnames

    def test_accepts_string_path(self, tmp_path):
        out = str(tmp_path / "str_path.csv")
        write_csv([], out)
        import pathlib

        assert pathlib.Path(out).exists()

    def test_csv_columns_match_fieldnames(self, tmp_path):
        """All _CSV_FIELDNAMES columns must appear in the header."""
        out = tmp_path / "cols.csv"
        write_csv([], out)
        content = out.read_text(encoding="utf-8")
        header_line = content.splitlines()[0]
        for field in bounces._CSV_FIELDNAMES:
            assert field in header_line, f"Missing column: {field}"


# ---------------------------------------------------------------------------
# scan_bounces MCP tool (Fix I2)
# ---------------------------------------------------------------------------

# Two canned bounce rows with known reason distribution:
#   2x "Invalid Recipient", 1x "Mailbox Full"
_TOOL_BOUNCE_ROWS = [
    {
        "first_name": "Bob",
        "last_name": "Jones",
        "email": "bob.jones@company.com",
        "reason": "Invalid Recipient",
        "date": "2025-03-15T10:30:00Z",
        "iso_date": "2025-03-15T10:30:00Z",
        "subject": "Undeliverable: Hello bob",
        "sender": "postmaster@mailserver.com",
        "body": "could not be delivered",
        "message_id": "AAA001",
        "has_attachments": False,
    },
    {
        "first_name": "Alice",
        "last_name": "Smith",
        "email": "alice.smith@example.com",
        "reason": "Invalid Recipient",
        "date": "2025-03-16T11:00:00Z",
        "iso_date": "2025-03-16T11:00:00Z",
        "subject": "Undeliverable: Hello alice",
        "sender": "postmaster@mailserver.com",
        "body": "could not be delivered",
        "message_id": "AAA002",
        "has_attachments": False,
    },
    {
        "first_name": "Carol",
        "last_name": "White",
        "email": "carol.white@example.com",
        "reason": "Mailbox Full",
        "date": "2025-03-17T12:00:00Z",
        "iso_date": "2025-03-17T12:00:00Z",
        "subject": "Delivery Failed",
        "sender": "mailer-daemon@example.com",
        "body": "Mailbox full",
        "message_id": "AAA003",
        "has_attachments": False,
    },
]


class TestScanBouncesTool:
    """Tests for the scan_bounces MCP tool in tools.py.

    The test suite imports via ``from src.microsoft_mcp import tools``, so the
    module lives under ``src.microsoft_mcp.tools`` in sys.modules.  Patch
    targets must use that prefix, not ``microsoft_mcp.tools``.

    Patch targets:
      - src.microsoft_mcp.tools._resolve_mail_folder  (folder resolution)
      - src.microsoft_mcp.tools._bounces.scan_folder  (folder scan)
      - src.microsoft_mcp.tools._bounces.write_csv    (CSV output)
    """

    def _get_tool_fn(self):
        from src.microsoft_mcp import tools

        return tools.scan_bounces.fn

    def test_returns_count_reasons_rows(self):
        """Returns dict with count, reasons map, and rows list."""
        with (
            mock.patch(
                "src.microsoft_mcp.tools._resolve_mail_folder",
                return_value="inbox",
            ) as mock_resolve,
            mock.patch(
                "src.microsoft_mcp.tools._bounces.scan_folder",
                return_value=_TOOL_BOUNCE_ROWS,
            ) as mock_scan,
        ):
            result = self._get_tool_fn()(folder="Inbox", limit=200)

        assert result["count"] == 3
        assert result["reasons"] == {"Invalid Recipient": 2, "Mailbox Full": 1}
        assert result["rows"] is _TOOL_BOUNCE_ROWS

        mock_resolve.assert_called_once_with("Inbox")
        mock_scan.assert_called_once()
        assert mock_scan.call_args[0][1] == "inbox"  # resolved folder_id
        assert mock_scan.call_args[1]["limit"] == 200

    def test_resolve_mail_folder_called_with_folder_arg(self):
        """_resolve_mail_folder is called with the raw folder argument."""
        with (
            mock.patch(
                "src.microsoft_mcp.tools._resolve_mail_folder",
                return_value="sentitems",
            ) as mock_resolve,
            mock.patch(
                "src.microsoft_mcp.tools._bounces.scan_folder",
                return_value=[],
            ),
        ):
            self._get_tool_fn()(folder="Sent", limit=50)

        mock_resolve.assert_called_once_with("Sent")

    def test_scan_folder_called_with_resolved_id_and_limit(self):
        """scan_folder receives the resolved folder_id and the limit kwarg."""
        with (
            mock.patch(
                "src.microsoft_mcp.tools._resolve_mail_folder",
                return_value="deleteditems",
            ),
            mock.patch(
                "src.microsoft_mcp.tools._bounces.scan_folder",
                return_value=[],
            ) as mock_scan,
        ):
            self._get_tool_fn()(folder="Deleted", limit=99)

        mock_scan.assert_called_once()
        assert mock_scan.call_args[0][1] == "deleteditems"
        assert mock_scan.call_args[1]["limit"] == 99

    def test_save_csv_calls_write_csv(self, tmp_path):
        """When save_csv is provided, write_csv is called with rows and path."""
        csv_path = str(tmp_path / "bounces.csv")
        with (
            mock.patch(
                "src.microsoft_mcp.tools._resolve_mail_folder",
                return_value="inbox",
            ),
            mock.patch(
                "src.microsoft_mcp.tools._bounces.scan_folder",
                return_value=_TOOL_BOUNCE_ROWS,
            ),
            mock.patch(
                "src.microsoft_mcp.tools._bounces.write_csv",
            ) as mock_write,
        ):
            result = self._get_tool_fn()(folder="Inbox", limit=200, save_csv=csv_path)

        mock_write.assert_called_once_with(_TOOL_BOUNCE_ROWS, csv_path)
        assert result["count"] == 3

    def test_no_save_csv_does_not_call_write_csv(self):
        """When save_csv is omitted, write_csv is NOT called."""
        with (
            mock.patch(
                "src.microsoft_mcp.tools._resolve_mail_folder",
                return_value="inbox",
            ),
            mock.patch(
                "src.microsoft_mcp.tools._bounces.scan_folder",
                return_value=_TOOL_BOUNCE_ROWS,
            ),
            mock.patch(
                "src.microsoft_mcp.tools._bounces.write_csv",
            ) as mock_write,
        ):
            self._get_tool_fn()(folder="Inbox", limit=200)

        mock_write.assert_not_called()

    def test_empty_folder_returns_zero_count(self):
        """When scan_folder returns no rows, count is 0 and reasons is empty."""
        with (
            mock.patch(
                "src.microsoft_mcp.tools._resolve_mail_folder",
                return_value="inbox",
            ),
            mock.patch(
                "src.microsoft_mcp.tools._bounces.scan_folder",
                return_value=[],
            ),
        ):
            result = self._get_tool_fn()(folder="Inbox", limit=200)

        assert result["count"] == 0
        assert result["reasons"] == {}
        assert result["rows"] == []
