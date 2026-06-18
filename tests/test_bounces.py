"""Tests for bounces.py — pure NDR classifier + DSN parser.

Import convention follows the repo pattern:
    from src.microsoft_mcp import <module>
"""

from src.microsoft_mcp import bounces
from src.microsoft_mcp.bounces import (
    classify_bounce_message,
    determine_bounce_reason,
    extract_email_from_text,
    is_bounce_message,
    parse_dsn_content,
    parse_name_from_email,
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
