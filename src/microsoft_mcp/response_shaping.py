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
