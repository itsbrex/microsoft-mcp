"""Helpers for Microsoft Graph To-Do task management."""

from datetime import date, timedelta
from typing import Any


def parse_due_date(text: str, *, today: date) -> dict[str, str]:
    """Parse due date text into Graph todoTask dueDateTime format.

    Accepts:
    - "today": current day
    - "tomorrow": next day
    - "+Nd": N days in future (e.g. "+3d")
    - "YYYY-MM-DD": absolute date

    Args:
        text: Due date text to parse
        today: Today's date (injected for deterministic testing)

    Returns:
        Dict with "dateTime" (YYYY-MM-DDT23:59:00) and "timeZone" (UTC)

    Raises:
        ValueError: If text cannot be parsed
    """
    text = text.strip()

    if text == "today":
        target_date = today
    elif text == "tomorrow":
        target_date = today + timedelta(days=1)
    elif text.startswith("+") and text.endswith("d"):
        try:
            days_offset = int(text[1:-1])
            target_date = today + timedelta(days=days_offset)
        except (ValueError, IndexError):
            raise ValueError(f"Invalid relative date format: {text}")
    else:
        try:
            target_date = date.fromisoformat(text)
        except ValueError:
            raise ValueError(f"Cannot parse due date: {text}")

    return {
        "dateTime": f"{target_date.isoformat()}T23:59:00",
        "timeZone": "UTC",
    }


def build_task_payload(
    *,
    title: str,
    importance: str = "normal",
    body: str | None = None,
    due: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Build Microsoft Graph todoTask payload.

    Args:
        title: Task title (required)
        importance: Importance level: "low", "normal", "high" (default: "normal")
        body: Task body/description (optional)
        due: Due date dict from parse_due_date (optional)

    Returns:
        Dict ready for Graph POST/PATCH to create/update task
    """
    payload: dict[str, Any] = {
        "title": title,
        "importance": importance,
    }

    if body is not None:
        payload["body"] = {
            "content": body,
            "contentType": "text",
        }

    if due is not None:
        payload["dueDateTime"] = due

    return payload


def build_linked_resource(web_url: str, display_name: str) -> dict[str, str]:
    """Build Microsoft Graph linkedResource for task.

    Args:
        web_url: URL to link
        display_name: Display name for the link

    Returns:
        Dict representing linkedResource
    """
    return {
        "applicationName": "Outlook",
        "webUrl": web_url,
        "displayName": display_name,
    }
