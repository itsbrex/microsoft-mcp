"""Email signature parser module.

Extracts contact information from email signatures and body content,
particularly useful for processing OOO (out-of-office) auto-replies.

Public API:
    normalize_phone_e164(phone, default_region="US") -> str
    parse_signature_block(text) -> dict
    parse_email_body(body, *, html=False, extract_alternatives=True) -> dict
"""

from __future__ import annotations

import re

# Signature block delimiters
_SIGNATURE_DELIMITERS = [
    r"^[\-_=]{3,}\s*$",
    r"^(Best|Kind|Warm)?\s*[Rr]egards?,?\s*$",
    r"^[Ss]incerely,?\s*$",
    r"^[Tt]hanks?,?\s*$",
    r"^[Cc]heers,?\s*$",
    r"^[Ww]ith\s+gratitude,?\s*$",
    r"^[Ss]ent\s+from\s+my\s+",
    r"^\*?CONFIDENTIALITY\s+NOTICE\*?",
    r"^Get\s+Outlook\s+for\s+",
]

_TITLE_PREFIXES = ["Mr.", "Ms.", "Mrs.", "Dr.", "Prof.", "Rev."]
_NAME_SUFFIXES = [
    "Jr.",
    "Sr.",
    "Jr",
    "Sr",
    "II",
    "III",
    "IV",
    "Esq.",
    "Esq",
    "Ph.D.",
    "M.D.",
    "CPA",
]

_JOB_TITLE_PATTERNS = [
    r"(?:Chief\s+)?(?:Executive|Financial|Operating|Technology|Marketing|Information|Revenue|Legal)\s+Officer",
    r"\bC[EFOTML]O\b",
    r"\bCEO\b",
    r"\bCFO\b",
    r"\bCOO\b",
    r"\bCTO\b",
    r"\bCMO\b",
    r"\bCIO\b",
    r"\bCRO\b",
    r"\bCLO\b",
    r"(?:Executive\s+)?(?:Vice\s+)?President",
    r"\b(?:E?VP|SVP|EVP)\b",
    r"(?:Executive\s+)?(?:Managing\s+)?(?:Senior\s+)?Director",
    r"(?:Executive\s+)?(?:Managing\s+)?Principal",
    r"(?:Senior\s+)?(?:General\s+)?Manager",
    r"(?:Senior\s+)?Partner",
    r"(?:Senior\s+)?(?:Staff\s+)?(?:Software\s+)?Engineer",
    r"(?:Senior\s+)?Analyst",
    r"(?:Senior\s+)?Associate",
    r"(?:Senior\s+)?Consultant",
    r"(?:Senior\s+)?Advisor",
    r"Coordinator",
    r"Specialist",
    r"Administrator",
    r"Supervisor",
    r"Lead",
    r"Head\s+of\s+\w+",
]

_PHONE_PATTERNS = [
    (r"(?:C|Cell|M|Mobile)[\s:]+([+\d\.\-\(\)\s]{10,20})", "mobile"),
    (r"(?:O|Office|W|Work|T|Tel|Phone)[\s:]+([+\d\.\-\(\)\s]{10,20})", "business"),
    (r"(?:H|Home)[\s:]+([+\d\.\-\(\)\s]{10,20})", "home"),
    (r"(?:F|Fax)[\s:]+([+\d\.\-\(\)\s]{10,20})", "fax"),
    (r"(?:Direct)[\s:]+([+\d\.\-\(\)\s]{10,20})", "business"),
]

_STANDALONE_PHONE_RE = re.compile(
    r"(?<![A-Za-z@])(\+?1?[\s\-\.]?\(?\d{3}\)?[\s\-\.]?\d{3}[\s\-\.]?\d{4})(?![A-Za-z@\d])"
)
_EMAIL_RE = re.compile(
    r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b", re.IGNORECASE
)
_URL_PATTERNS = [
    re.compile(r"https?://[^\s<>\"'\)]+"),
    re.compile(r"www\.[^\s<>\"'\)]+\.[a-z]{2,}[^\s<>\"'\)]*", re.IGNORECASE),
]
_LINKEDIN_RE = re.compile(
    r"(?:https?://)?(?:www\.)?linkedin\.com/in/([a-zA-Z0-9_-]+)", re.IGNORECASE
)
_TWITTER_RE = re.compile(r"(?<![A-Za-z0-9._%+-])@([a-zA-Z0-9_]{1,15})(?![A-Za-z0-9@])")

_ALT_CONTACT_PATTERNS = [
    r"(?:Ms\.|Mr\.|Mrs\.|Dr\.)?\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\s*[-–:]\s*([A-Za-z\s]+?):\s*(\S+@\S+)",
    r"(?:contact|reach\s+out\s+to|email)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\s+(?:at\s+)?[\(<]?(\S+@\S+)[\)>]?",
    r"([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+):\s*([A-Za-z\s]+?)\s+(\S+@\S+)",
    r"([A-Za-z\s]+?):\s*(\S+@\S+)\s*(?:\+|and)?\s*([\d\.\-\(\)\s]{10,20})?",
]

_JOB_CHANGE_PATTERNS = [
    (r"(?:I'm|I\s+am)\s+no\s+longer\s+(?:at|with)\s+(.+?)(?:\.|,|$)", "left_company"),
    (r"(?:left|departed|resigned\s+from)\s+(.+?)(?:\.|,|$)", "left_company"),
    (r"(?:now\s+at|joined|moved\s+to)\s+(.+?)(?:\.|,|$)", "new_company"),
    (r"(?:new\s+email|new\s+address|reach\s+me\s+at):\s*(\S+@\S+)", "new_email"),
]

_PERSONAL_EMAIL_DOMAINS = {
    "gmail.com",
    "yahoo.com",
    "outlook.com",
    "hotmail.com",
    "icloud.com",
    "me.com",
    "mac.com",
    "aol.com",
    "proton.me",
    "protonmail.com",
    "live.com",
    "msn.com",
}

# Maximum number of characters the regex pipeline operates on.  Attacker-
# controlled bodies can be arbitrarily large; several patterns (alt-contact,
# job-change, company) show super-linear CPU growth beyond this size.
_MAX_BODY_LEN = 20_000


# ---------------------------------------------------------------------------
# Phone Normalization
# ---------------------------------------------------------------------------


def normalize_phone_e164(phone: str, default_region: str = "US") -> str:  # noqa: ARG001
    """Normalize a phone number to E.164 format.

    Rules (US default):
    - 10 digits → +1XXXXXXXXXX
    - 11 digits starting with 1 → +1XXXXXXXXXX
    - Already starts with + → keep + and digits
    - <10 digits or junk → ""

    Examples::

        >>> normalize_phone_e164("(949) 462-4106")
        '+19494624106'
        >>> normalize_phone_e164("+1-949-462-4106")
        '+19494624106'
        >>> normalize_phone_e164("junk")
        ''
    """
    if not phone:
        return ""
    phone = phone.strip()
    if phone.startswith("+"):
        digits = re.sub(r"\D", "", phone[1:])
        return f"+{digits}" if digits else ""
    digits = re.sub(r"\D", "", phone)
    if not digits:
        return ""
    if digits.startswith("011") and len(digits) > 3:
        return f"+{digits[3:]}"
    if len(digits) == 10:
        return f"+1{digits}"
    if len(digits) == 11 and digits.startswith("1"):
        return f"+{digits}"
    if 11 <= len(digits) <= 15:
        return f"+{digits}"
    return ""


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _extract_emails(text: str) -> list[tuple[str, str]]:
    results: list[tuple[str, str]] = []
    seen: set[str] = set()
    for m in _EMAIL_RE.finditer(text):
        addr = m.group(0).lower()
        if addr in seen:
            continue
        seen.add(addr)
        domain = addr.split("@")[-1]
        results.append((addr, "home" if domain in _PERSONAL_EMAIL_DOMAINS else "work"))
    return results


def _extract_phones(text: str) -> list[tuple[str, str]]:
    results: list[tuple[str, str]] = []
    seen: set[str] = set()
    for pattern, phone_type in _PHONE_PATTERNS:
        for m in re.finditer(pattern, text, re.IGNORECASE):
            n = normalize_phone_e164(m.group(1))
            if n and n not in seen:
                seen.add(n)
                results.append((n, phone_type))
    for m in _STANDALONE_PHONE_RE.finditer(text):
        n = normalize_phone_e164(m.group(1))
        if n and n not in seen:
            seen.add(n)
            results.append((n, "other"))
    return results


def _extract_urls(text: str) -> list[str]:
    urls: list[str] = []
    seen: set[str] = set()
    for pat in _URL_PATTERNS:
        for m in pat.finditer(text):
            url = m.group(0).rstrip(".,;:)")
            if url.startswith("www."):
                url = f"https://{url}"
            # LinkedIn URLs are captured separately via _extract_linkedin
            if "linkedin.com" in url.lower():
                continue
            if url.lower() not in seen:
                seen.add(url.lower())
                urls.append(url)
    return urls


def _extract_linkedin(text: str) -> str:
    m = _LINKEDIN_RE.search(text)
    return f"https://linkedin.com/in/{m.group(1)}" if m else ""


def _extract_twitter(text: str) -> str:
    for m in _TWITTER_RE.finditer(text):
        if "@" not in text[max(0, m.start() - 20) : m.start()]:
            return f"@{m.group(1)}"
    return ""


def _parse_name(text: str) -> tuple[str, str, str, str, str, str]:
    """Return (prefix, first, middle, last, suffix, full_name)."""
    if not text:
        return ("", "", "", "", "", "")
    text = text.strip()
    full_name = text
    prefix = ""
    for p in _TITLE_PREFIXES:
        if text.startswith(p):
            prefix = p
            text = text[len(p) :].strip()
            break
    suffix = ""
    for s in _NAME_SUFFIXES:
        for pat in [f", {s}$", f",{s}$", f" {s}$"]:
            mm = re.search(pat, text, re.IGNORECASE)
            if mm:
                suffix = s
                text = text[: mm.start()].strip()
                break
        if suffix:
            break
    parts = text.split()
    if not parts:
        return (prefix, "", "", "", suffix, full_name)
    if len(parts) == 1:
        return (prefix, parts[0], "", "", suffix, full_name)
    if len(parts) == 2:
        return (prefix, parts[0], "", parts[1], suffix, full_name)
    return (prefix, parts[0], " ".join(parts[1:-1]), parts[-1], suffix, full_name)


def _parse_name_from_email(email: str) -> tuple[str, str]:
    if not email or "@" not in email:
        return ("", "")
    local = re.sub(r"\d+", "", email.split("@")[0])
    parts = re.split(r"[._-]", local)
    if len(parts) >= 2:
        return (parts[0].capitalize(), parts[-1].capitalize())
    camel = re.findall(r"[A-Z]?[a-z]+|[A-Z]+(?=[A-Z][a-z]|\b)", parts[0])
    if len(camel) >= 2:
        return (camel[0].capitalize(), camel[-1].capitalize())
    return (parts[0].capitalize(), "")


def _extract_job_title(text: str) -> str:
    for pattern in _JOB_TITLE_PATTERNS:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            start = text.rfind("\n", 0, m.start()) + 1
            end = text.find("\n", m.end())
            if end == -1:
                end = len(text)
            line = re.sub(r"^\s*[-•*]\s*", "", text[start:end].strip())
            line = re.sub(r"\s+", " ", line)
            if len(line) < 100:
                return line
    return ""


def _extract_company(text: str, emails: list[tuple[str, str]] | None = None) -> str:
    # Require at least 3 chars of company name before a recognizable legal suffix.
    # Use \b on the suffix to avoid "Co." matching inside longer words like "Commercial".
    for pattern in [
        r"([A-Z][A-Za-z\s&]{2,}?\b(?:Inc\.?|LLC|Corp\.?|Ltd\.?|Company|Group|Partners?))\b",
        r"([A-Z][A-Za-z\s&]{2,}?\b(?:International|Enterprises?|Solutions?|Services?|Technologies?|Consulting))\b",
    ]:
        for m in re.finditer(pattern, text):
            company = m.group(1).strip()
            # Skip short noise like "Co" or entries that are clearly job titles
            if 5 <= len(company) < 80 and not re.search(
                r"\b(?:Senior|Director|Manager|Engineer|Analyst|Associate|Coordinator)\b",
                company,
                re.IGNORECASE,
            ):
                return company
    if emails:
        for addr, email_type in emails:
            if email_type == "work":
                name_part = addr.split("@")[-1].split(".")[0]
                if name_part and len(name_part) > 2:
                    return name_part.replace("-", " ").title()
    return ""


def _calculate_confidence(contact: dict) -> float:
    score = 0.0
    if contact.get("first_name") and contact.get("last_name"):
        score += 0.2
    if contact.get("work_email"):
        score += 0.2
        if contact.get("company"):
            domain = contact["work_email"].split("@")[-1].lower()
            if any(
                w in domain for w in contact["company"].lower().split() if len(w) > 3
            ):
                score += 0.05
    if contact.get("job_title"):
        score += 0.15
    if contact.get("company"):
        score += 0.15
    if contact.get("business_phone") or contact.get("mobile_phone"):
        score += 0.1
    if contact.get("website"):
        score += 0.05
    if contact.get("linkedin") or contact.get("twitter"):
        score += 0.05
    return min(score, 1.0)


def _find_signature_start(text: str) -> int:
    lines = text.split("\n")
    best_pos = -1
    for i, line in enumerate(lines):
        for pattern in _SIGNATURE_DELIMITERS:
            if re.match(pattern, line.strip(), re.IGNORECASE):
                pos = sum(len(lines[j]) + 1 for j in range(i))
                if best_pos == -1 or pos < best_pos:
                    best_pos = pos
                break
    return best_pos


def _extract_signature_block(text: str) -> str:
    start = _find_signature_start(text)
    if start == -1:
        lines = text.strip().split("\n")
        return "\n".join(lines[-15:]) if len(lines) > 5 else text
    return text[start:]


def _empty_contact() -> dict:
    return {
        "first_name": "",
        "last_name": "",
        "full_name": "",
        "job_title": "",
        "company": "",
        "work_email": "",
        "mobile_phone": "",
        "business_phone": "",
        "website": "",
        "linkedin": "",
        "twitter": "",
        "confidence_score": 0.0,
    }


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def parse_signature_block(text: str) -> dict:
    """Parse a plain-text signature block and return a contact dict.

    Keys: first_name, last_name, full_name, job_title, company,
    work_email, mobile_phone, business_phone, website, linkedin, twitter,
    confidence_score.
    """
    contact = _empty_contact()
    emails = _extract_emails(text)
    for addr, email_type in emails:
        if email_type == "work" and not contact["work_email"]:
            contact["work_email"] = addr
            break

    phones = _extract_phones(text)
    for number, phone_type in phones:
        if phone_type == "mobile" and not contact["mobile_phone"]:
            contact["mobile_phone"] = number
        elif phone_type in ("business", "other") and not contact["business_phone"]:
            contact["business_phone"] = number

    urls = _extract_urls(text)
    if urls:
        contact["website"] = urls[0]

    contact["linkedin"] = _extract_linkedin(text)
    contact["twitter"] = _extract_twitter(text)
    contact["job_title"] = _extract_job_title(text)
    contact["company"] = _extract_company(text, emails)

    for line in text.strip().split("\n")[:5]:
        line = line.strip()
        if not line:
            continue
        if "@" in line or re.match(r"^https?://", line) or re.match(r"^www\.", line):
            continue
        if re.match(r"^[\d\(\+]", line):
            continue
        if re.match(
            r"^(?:Mr\.|Ms\.|Mrs\.|Dr\.)?\s*[A-Z][a-z]+(?:\s+[A-Z]\.?\s*)?(?:\s+[A-Z][a-z]+)+",
            line,
        ):
            _, first, _, last, _, full = _parse_name(line)
            contact["first_name"] = first
            contact["last_name"] = last
            contact["full_name"] = full
            break

    if not contact["first_name"] and contact["work_email"]:
        first, last = _parse_name_from_email(contact["work_email"])
        contact["first_name"] = first
        contact["last_name"] = last

    contact["confidence_score"] = _calculate_confidence(contact)
    return contact


def parse_email_body(
    body: str,
    *,
    html: bool = False,
    extract_alternatives: bool = True,
) -> dict:
    """Parse email body to extract contacts and job-change signals.

    Args:
        body: Email body text (or HTML when html=True).
        html: Strip HTML tags before parsing.
        extract_alternatives: Extract alternative contacts from OOO prose.

    Returns:
        Dict with keys:
          - contacts: list of contact dicts (same schema as parse_signature_block)
          - job_changes: dict with optional keys left_company, new_company, new_email
    """
    if html:
        from .response_shaping import _html_to_text  # lazy import; keeps module pure

        body = _html_to_text(body)

    # Cap body length AFTER any HTML-to-text conversion so the regex pipeline
    # never operates on arbitrarily large attacker-controlled input.
    body = body[:_MAX_BODY_LEN]

    contacts: list[dict] = []

    sig_block = _extract_signature_block(body)
    if sig_block:
        primary = parse_signature_block(sig_block)
        if primary.get("work_email") or (
            primary.get("first_name") and primary.get("last_name")
        ):
            contacts.append(primary)

    if extract_alternatives:
        seen_emails: set[str] = {
            c["work_email"] for c in contacts if c.get("work_email")
        }
        for pattern in _ALT_CONTACT_PATTERNS:
            for m in re.finditer(pattern, body, re.IGNORECASE):
                groups = m.groups()
                if not groups:
                    continue
                alt = _empty_contact()
                name_or_dept = groups[0].strip()
                if re.match(r"^[A-Z][a-z]+(?:\s+[A-Z][a-z]+)+$", name_or_dept):
                    _, first, _, last, _, full = _parse_name(name_or_dept)
                    alt["first_name"] = first
                    alt["last_name"] = last
                    alt["full_name"] = full
                else:
                    alt["job_title"] = name_or_dept
                for grp in groups[1:]:
                    if not grp:
                        continue
                    grp = grp.strip()
                    if "@" in grp:
                        email = grp.lower()
                        if email not in seen_emails:
                            seen_emails.add(email)
                            alt["work_email"] = email
                    elif not alt["job_title"]:
                        alt["job_title"] = grp
                    else:
                        phone = normalize_phone_e164(grp)
                        if phone and not alt["business_phone"]:
                            alt["business_phone"] = phone
                if alt.get("work_email") or (
                    alt.get("first_name") and alt.get("last_name")
                ):
                    if not alt["first_name"] and alt.get("work_email"):
                        first, last = _parse_name_from_email(alt["work_email"])
                        alt["first_name"] = first
                        alt["last_name"] = last
                    alt["confidence_score"] = _calculate_confidence(alt)
                    contacts.append(alt)

    job_changes: dict[str, str] = {}
    for pattern, intel_type in _JOB_CHANGE_PATTERNS:
        mm = re.search(pattern, body, re.IGNORECASE)
        if mm:
            job_changes[intel_type] = mm.group(1).strip()

    return {"contacts": contacts, "job_changes": job_changes}
