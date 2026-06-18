"""Pure, Graph-free helpers for Outlook inbox message rules.

All network I/O lives in tools.py; this module only builds payloads,
summarizes rules for display, and converts to/from YAML templates.
"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any

RULE_LIST_FIELDS = "id,displayName,sequence,isEnabled,conditions,actions"
RULE_DETAIL_FIELDS = (
    "id,displayName,sequence,isEnabled,hasError,isReadOnly,"
    "conditions,actions,exceptions"
)

# Human-readable labels for the boolean/scalar predicate keys.
_CONDITION_LABELS: dict[str, str] = {
    "hasAttachments": "has attachments",
    "isApprovalRequest": "is approval request",
    "isAutomaticForward": "is auto-forward",
    "isAutomaticReply": "is auto-reply",
    "isEncrypted": "is encrypted",
    "isMeetingRequest": "is meeting request",
    "isMeetingResponse": "is meeting response",
    "isNonDeliveryReport": "is NDR/bounce",
    "isPermissionControlled": "is permission-controlled",
    "isReadReceipt": "is read receipt",
    "isSigned": "is signed",
    "isVoicemail": "is voicemail",
    "sentToMe": "sent to me",
    "sentCcMe": "cc's me",
    "sentOnlyToMe": "sent only to me",
    "sentToOrCcMe": "sent to or cc's me",
    "notSentToMe": "not sent to me",
}
_LIST_CONDITION_LABELS: dict[str, str] = {
    "bodyContains": "body contains",
    "bodyOrSubjectContains": "body/subject contains",
    "headerContains": "header contains",
    "subjectContains": "subject contains",
    "senderContains": "sender contains",
    "recipientContains": "recipient contains",
    "categories": "categorized",
}


def _addrs(recips: list[dict[str, Any]]) -> list[str]:
    out = []
    for r in recips or []:
        ea = r.get("emailAddress", {}) if isinstance(r, dict) else {}
        out.append(ea.get("address") or ea.get("name") or "")
    return [a for a in out if a]


def summarize_conditions(conditions: dict[str, Any] | None) -> str:
    if not conditions:
        return "(any message)"
    parts: list[str] = []
    for key, label in _LIST_CONDITION_LABELS.items():
        vals = conditions.get(key)
        if vals:
            parts.append(f"{label} {', '.join(vals)}")
    if conditions.get("fromAddresses"):
        parts.append(f"from {', '.join(_addrs(conditions['fromAddresses']))}")
    if conditions.get("sentToAddresses"):
        parts.append(f"sent to {', '.join(_addrs(conditions['sentToAddresses']))}")
    if conditions.get("importance"):
        parts.append(f"importance={conditions['importance']}")
    for key, label in _CONDITION_LABELS.items():
        if conditions.get(key):
            parts.append(label)
    return "; ".join(parts) if parts else "(any message)"


def summarize_actions(actions: dict[str, Any] | None) -> str:
    if not actions:
        return "(no actions)"
    parts: list[str] = []
    if actions.get("moveToFolder"):
        parts.append(f"move to {actions['moveToFolder']}")
    if actions.get("copyToFolder"):
        parts.append(f"copy to {actions['copyToFolder']}")
    if actions.get("assignCategories"):
        parts.append(f"categorize {', '.join(actions['assignCategories'])}")
    if actions.get("markImportance"):
        parts.append(f"mark importance {actions['markImportance']}")
    if actions.get("markAsRead"):
        parts.append("mark as read")
    if actions.get("forwardTo"):
        parts.append(f"forward to {', '.join(_addrs(actions['forwardTo']))}")
    if actions.get("redirectTo"):
        parts.append(f"redirect to {', '.join(_addrs(actions['redirectTo']))}")
    if actions.get("delete"):
        parts.append("delete")
    if actions.get("permanentDelete"):
        parts.append("permanently delete")
    if actions.get("stopProcessingRules"):
        parts.append("stop processing further rules")
    return "; ".join(parts) if parts else "(no actions)"


def shape_rule_summary(rule: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": rule.get("id"),
        "display_name": rule.get("displayName"),
        "sequence": rule.get("sequence"),
        "is_enabled": rule.get("isEnabled"),
        "conditions_summary": summarize_conditions(rule.get("conditions")),
        "actions_summary": summarize_actions(rule.get("actions")),
    }


def _recipients(emails: list[str] | None) -> list[dict[str, Any]] | None:
    if not emails:
        return None
    return [{"emailAddress": {"address": e}} for e in emails]


def build_rule_payload(
    *,
    display_name: str,
    sequence: int = 1,
    is_enabled: bool = True,
    sender_contains: list[str] | None = None,
    subject_contains: list[str] | None = None,
    body_contains: list[str] | None = None,
    from_addresses: list[str] | None = None,
    has_attachments: bool | None = None,
    importance: str | None = None,
    move_to_folder: str | None = None,
    copy_to_folder: str | None = None,
    assign_categories: list[str] | None = None,
    mark_as_read: bool | None = None,
    mark_importance: str | None = None,
    forward_to: list[str] | None = None,
    delete: bool | None = None,
    stop_processing_rules: bool | None = None,
) -> dict[str, Any]:
    conditions: dict[str, Any] = {}
    if sender_contains:
        conditions["senderContains"] = sender_contains
    if subject_contains:
        conditions["subjectContains"] = subject_contains
    if body_contains:
        conditions["bodyContains"] = body_contains
    if from_addresses:
        conditions["fromAddresses"] = _recipients(from_addresses)
    if has_attachments is not None:
        conditions["hasAttachments"] = has_attachments
    if importance:
        conditions["importance"] = importance

    actions: dict[str, Any] = {}
    if move_to_folder:
        actions["moveToFolder"] = move_to_folder
    if copy_to_folder:
        actions["copyToFolder"] = copy_to_folder
    if assign_categories:
        actions["assignCategories"] = assign_categories
    if mark_as_read is not None:
        actions["markAsRead"] = mark_as_read
    if mark_importance:
        actions["markImportance"] = mark_importance
    if forward_to:
        actions["forwardTo"] = _recipients(forward_to)
    if delete is not None:
        actions["delete"] = delete
    if stop_processing_rules is not None:
        actions["stopProcessingRules"] = stop_processing_rules

    payload: dict[str, Any] = {
        "displayName": display_name,
        "sequence": sequence,
        "isEnabled": is_enabled,
    }
    if conditions:
        payload["conditions"] = conditions
    if actions:
        payload["actions"] = actions
    return payload


# ---------------------------------------------------------------------------
# YAML template ⇄ Graph payload converters
# ---------------------------------------------------------------------------

_IMPORTANCE_VALUES = {"low", "normal", "high"}


def template_to_rule_payload(
    tpl: dict[str, Any],
    folder_resolver: Callable[[str], str] | None = None,
) -> dict[str, Any]:
    """Convert a snake_case YAML template dict to a Graph rule payload.

    Delegates to :func:`build_rule_payload` so all camel-case mapping lives
    in one place.  If *folder_resolver* is provided, ``move_to`` / ``copy_to``
    folder names are resolved to folder IDs before being passed through.
    """
    conds: dict[str, Any] = tpl.get("conditions") or {}
    acts: dict[str, Any] = tpl.get("actions") or {}

    def _resolve(name: str) -> str:
        return folder_resolver(name) if folder_resolver else name

    move_to = acts.get("move_to")
    copy_to = acts.get("copy_to")

    return build_rule_payload(
        display_name=tpl["name"],
        sequence=tpl.get("sequence", 1),
        is_enabled=tpl.get("enabled", True),
        # conditions
        sender_contains=conds.get("sender_contains"),
        subject_contains=conds.get("subject_contains"),
        body_contains=conds.get("body_contains"),
        from_addresses=conds.get("from_addresses"),
        has_attachments=conds.get("has_attachments"),
        importance=conds.get("importance"),
        # actions
        move_to_folder=_resolve(move_to) if move_to else None,
        copy_to_folder=_resolve(copy_to) if copy_to else None,
        assign_categories=acts.get("assign_categories"),
        mark_as_read=acts.get("mark_as_read"),
        mark_importance=acts.get("mark_importance"),
        forward_to=acts.get("forward_to"),
        delete=acts.get("delete"),
        stop_processing_rules=acts.get("stop_processing"),
    )


def rule_to_template(
    rule: dict[str, Any],
    folder_namer: Callable[[str], str] | None = None,
) -> dict[str, Any]:
    """Convert a Graph message rule dict to a snake_case template dict.

    Inverse of :func:`template_to_rule_payload`.  If *folder_namer* is
    provided, folder IDs in ``moveToFolder`` / ``copyToFolder`` are resolved
    to human-readable names.
    """
    conds_raw: dict[str, Any] = rule.get("conditions") or {}
    acts_raw: dict[str, Any] = rule.get("actions") or {}

    def _name(fid: str) -> str:
        return folder_namer(fid) if folder_namer else fid

    conds: dict[str, Any] = {}
    if conds_raw.get("senderContains"):
        conds["sender_contains"] = conds_raw["senderContains"]
    if conds_raw.get("subjectContains"):
        conds["subject_contains"] = conds_raw["subjectContains"]
    if conds_raw.get("bodyContains"):
        conds["body_contains"] = conds_raw["bodyContains"]
    if conds_raw.get("fromAddresses"):
        conds["from_addresses"] = [
            (r.get("emailAddress", {}).get("address") or "")
            for r in conds_raw["fromAddresses"]
        ]
    if conds_raw.get("hasAttachments") is not None:
        conds["has_attachments"] = conds_raw["hasAttachments"]
    if conds_raw.get("importance"):
        conds["importance"] = conds_raw["importance"]

    acts: dict[str, Any] = {}
    if acts_raw.get("moveToFolder"):
        acts["move_to"] = _name(acts_raw["moveToFolder"])
    if acts_raw.get("copyToFolder"):
        acts["copy_to"] = _name(acts_raw["copyToFolder"])
    if acts_raw.get("assignCategories"):
        acts["assign_categories"] = acts_raw["assignCategories"]
    if acts_raw.get("markAsRead") is not None:
        acts["mark_as_read"] = acts_raw["markAsRead"]
    if acts_raw.get("markImportance"):
        acts["mark_importance"] = acts_raw["markImportance"]
    if acts_raw.get("forwardTo"):
        acts["forward_to"] = [
            (r.get("emailAddress", {}).get("address") or "")
            for r in acts_raw["forwardTo"]
        ]
    if acts_raw.get("delete") is not None:
        acts["delete"] = acts_raw["delete"]
    if acts_raw.get("stopProcessingRules") is not None:
        acts["stop_processing"] = acts_raw["stopProcessingRules"]

    tpl: dict[str, Any] = {
        "name": rule.get("displayName", ""),
        "enabled": rule.get("isEnabled", True),
    }
    if rule.get("sequence") is not None:
        tpl["sequence"] = rule["sequence"]
    if conds:
        tpl["conditions"] = conds
    if acts:
        tpl["actions"] = acts
    return tpl


def validate_template(tpl: dict[str, Any]) -> list[str]:
    """Validate a snake_case rule template dict.

    Returns a list of human-readable error strings; an empty list means valid.
    """
    errors: list[str] = []

    if not tpl.get("name"):
        errors.append("'name' is required")

    conds: dict[str, Any] = tpl.get("conditions") or {}
    acts: dict[str, Any] = tpl.get("actions") or {}

    # At least one condition key with a non-empty/non-false value
    has_condition = any(
        (v is not False and v is not None and v != [] and v != {})
        for v in conds.values()
    )
    if not has_condition:
        errors.append("at least one condition is required")

    # At least one real action OR stop_processing: true
    stop = acts.get("stop_processing")
    has_action = stop is True or any(
        k != "stop_processing"
        and v is not None
        and v is not False
        and v != []
        and v != {}
        for k, v in acts.items()
    )
    if not has_action:
        errors.append("at least one action (or stop_processing: true) is required")

    if "enabled" in tpl and not isinstance(tpl["enabled"], bool):
        errors.append("'enabled' must be a boolean")

    if "sequence" in tpl and not isinstance(tpl["sequence"], int):
        errors.append("'sequence' must be an integer")

    # importance in conditions
    cond_imp = conds.get("importance")
    if cond_imp is not None and str(cond_imp).lower() not in _IMPORTANCE_VALUES:
        errors.append(
            f"conditions.importance must be one of {sorted(_IMPORTANCE_VALUES)}, got {cond_imp!r}"
        )

    # mark_importance in actions
    act_imp = acts.get("mark_importance")
    if act_imp is not None and str(act_imp).lower() not in _IMPORTANCE_VALUES:
        errors.append(
            f"actions.mark_importance must be one of {sorted(_IMPORTANCE_VALUES)}, got {act_imp!r}"
        )

    return errors
