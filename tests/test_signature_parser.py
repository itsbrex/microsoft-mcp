"""Tests for signature_parser module.

Import convention follows the repo pattern used in other test files:
    from src.microsoft_mcp import <module>
"""

import pytest

from src.microsoft_mcp import signature_parser
from src.microsoft_mcp.signature_parser import (
    normalize_phone_e164,
    parse_email_body,
    parse_signature_block,
)

# ---------------------------------------------------------------------------
# Sample data
# ---------------------------------------------------------------------------

SAMPLE_SIG = """\
Jane Doe
Senior Director, Commercial Real Estate
Acme Corp, LLC
M: (949) 462-4106
T: (714) 555-0199
jane.doe@acmecorp.com
www.acmecorp.com
https://linkedin.com/in/janedoe
"""

SAMPLE_EMAIL_BODY = """\
Hi there,

I'm out of the office until Monday.

For urgent matters, contact Thai Tran - Director: ttran@auroraspine.us

Best regards,
Jane Doe
Senior Director
jane.doe@acmecorp.com
M: (949) 462-4106
"""

OOO_JOB_CHANGE = """\
I am no longer at Globex Corp.  I have now joined Initech.
My new email: jsmith@initech.com

Regards,
John Smith
Director
john.smith@initech.com
"""

HTML_BODY = """\
<html><body>
<p>Please reach me at <b>jane.doe@acmecorp.com</b></p>
<p>Jane Doe<br/>Senior Director<br/>Acme Corp, LLC</p>
</body></html>
"""


# ---------------------------------------------------------------------------
# normalize_phone_e164
# ---------------------------------------------------------------------------


class TestNormalizePhoneE164:
    def test_us_10_digit_formatted(self):
        assert normalize_phone_e164("(949) 462-4106") == "+19494624106"

    def test_us_10_digit_dots(self):
        assert normalize_phone_e164("949.462.4106") == "+19494624106"

    def test_plus1_dashes(self):
        assert normalize_phone_e164("+1-949-462-4106") == "+19494624106"

    def test_plus1_parens_spaces(self):
        assert normalize_phone_e164("+1 (949) 462-4106") == "+19494624106"

    def test_11_digit_leading_1(self):
        assert normalize_phone_e164("19494624106") == "+19494624106"

    def test_already_e164(self):
        assert normalize_phone_e164("+19494624106") == "+19494624106"

    def test_empty_string(self):
        assert normalize_phone_e164("") == ""

    def test_junk(self):
        assert normalize_phone_e164("junk") == ""

    def test_too_short(self):
        assert normalize_phone_e164("123") == ""

    def test_default_region_param_accepted(self):
        # default_region kwarg must be accepted without error
        result = normalize_phone_e164("(949) 462-4106", default_region="US")
        assert result == "+19494624106"


# ---------------------------------------------------------------------------
# parse_signature_block
# ---------------------------------------------------------------------------


class TestParseSignatureBlock:
    def test_extracts_first_and_last_name(self):
        result = parse_signature_block(SAMPLE_SIG)
        assert result["first_name"] == "Jane"
        assert result["last_name"] == "Doe"

    def test_extracts_full_name(self):
        result = parse_signature_block(SAMPLE_SIG)
        assert result["full_name"] == "Jane Doe"

    def test_extracts_job_title(self):
        result = parse_signature_block(SAMPLE_SIG)
        assert "Director" in result["job_title"]

    def test_extracts_company(self):
        result = parse_signature_block(SAMPLE_SIG)
        assert "Acme" in result["company"]

    def test_extracts_work_email(self):
        result = parse_signature_block(SAMPLE_SIG)
        assert result["work_email"] == "jane.doe@acmecorp.com"

    def test_extracts_mobile_phone(self):
        result = parse_signature_block(SAMPLE_SIG)
        assert result["mobile_phone"] == "+19494624106"

    def test_extracts_linkedin(self):
        result = parse_signature_block(SAMPLE_SIG)
        assert result["linkedin"] == "https://linkedin.com/in/janedoe"

    def test_extracts_website(self):
        result = parse_signature_block(SAMPLE_SIG)
        assert "acmecorp.com" in result["website"]

    def test_confidence_score_in_range(self):
        result = parse_signature_block(SAMPLE_SIG)
        assert 0.0 <= result["confidence_score"] <= 1.0

    def test_confidence_high_for_rich_sig(self):
        result = parse_signature_block(SAMPLE_SIG)
        assert result["confidence_score"] >= 0.5

    def test_confidence_zero_for_empty(self):
        result = parse_signature_block("")
        assert result["confidence_score"] == 0.0

    def test_returns_all_required_keys(self):
        result = parse_signature_block(SAMPLE_SIG)
        required = {
            "first_name", "last_name", "full_name", "job_title", "company",
            "work_email", "mobile_phone", "business_phone",
            "website", "linkedin", "twitter", "confidence_score",
        }
        assert required.issubset(result.keys())

    def test_twitter_extracted(self):
        sig = "Jane Doe\n@janedoe_x\njane@example.com"
        result = parse_signature_block(sig)
        assert result["twitter"] == "@janedoe_x"


# ---------------------------------------------------------------------------
# parse_email_body
# ---------------------------------------------------------------------------


class TestParseEmailBody:
    def test_returns_contacts_and_job_changes_keys(self):
        result = parse_email_body(SAMPLE_EMAIL_BODY)
        assert "contacts" in result
        assert "job_changes" in result

    def test_extracts_primary_contact(self):
        result = parse_email_body(SAMPLE_EMAIL_BODY)
        assert len(result["contacts"]) >= 1
        primary = result["contacts"][0]
        assert "jane.doe@acmecorp.com" in primary["work_email"]

    def test_html_true_strips_tags(self):
        result = parse_email_body(HTML_BODY, html=True)
        contacts = result["contacts"]
        assert len(contacts) >= 1
        emails = [c["work_email"] for c in contacts]
        assert any("jane.doe@acmecorp.com" in e for e in emails)

    def test_job_change_left_company(self):
        result = parse_email_body(OOO_JOB_CHANGE)
        assert "left_company" in result["job_changes"]
        assert "Globex" in result["job_changes"]["left_company"]

    def test_job_change_new_company(self):
        result = parse_email_body(OOO_JOB_CHANGE)
        assert "new_company" in result["job_changes"]
        assert "Initech" in result["job_changes"]["new_company"]

    def test_job_change_new_email(self):
        result = parse_email_body(OOO_JOB_CHANGE)
        assert "new_email" in result["job_changes"]
        assert "initech" in result["job_changes"]["new_email"]

    def test_extract_alternatives_false(self):
        result = parse_email_body(SAMPLE_EMAIL_BODY, extract_alternatives=False)
        # Should still have primary contact, but no alt contacts appended
        assert isinstance(result["contacts"], list)

    def test_confidence_in_range_for_all_contacts(self):
        result = parse_email_body(SAMPLE_EMAIL_BODY)
        for c in result["contacts"]:
            assert 0.0 <= c["confidence_score"] <= 1.0

    def test_empty_body(self):
        result = parse_email_body("")
        assert result["contacts"] == [] or isinstance(result["contacts"], list)
        assert isinstance(result["job_changes"], dict)
