from enum import Enum
from dataclasses import dataclass


class ResponseProfile(str, Enum):
    RAW = "raw"
    DETAIL = "detail"
    SUMMARY = "summary"

    @classmethod
    def default_for_operation(cls, operation: str) -> "ResponseProfile":
        return cls.SUMMARY if operation in {"list", "search"} else cls.DETAIL


@dataclass(frozen=True)
class BudgetHints:
    include_body: bool
    max_items: int

    @classmethod
    def for_operation(cls, tool_name: str) -> "BudgetHints":
        return cls(include_body=False, max_items=25)


# ---------------------------------------------------------------------------
# Global Graph payload cleanup
# ---------------------------------------------------------------------------

ODATA_KEYS = {"@odata.context", "@odata.etag", "@odata.type", "@odata.id", "@odata.count"}
NOISE_KEYS = {
    "changeKey",
    "parentFolderId",
    "calendar@odata.associationLink",
    "calendar@odata.navigationLink",
}
_EMPTY = (None, "", [], {})


def cleanup_graph_payload(value: object) -> object:
    if isinstance(value, dict):
        cleaned: dict = {}
        for key, child in value.items():
            if key in ODATA_KEYS or key in NOISE_KEYS:
                continue
            next_value = cleanup_graph_payload(child)
            if next_value in _EMPTY:
                continue
            cleaned[key] = next_value
        return cleaned
    if isinstance(value, list):
        return [
            item
            for item in (cleanup_graph_payload(v) for v in value)
            if item not in _EMPTY
        ]
    return value
