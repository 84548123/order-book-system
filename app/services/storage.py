from __future__ import annotations

import base64
from typing import Any

import httpx

from app.config import settings


class StorageError(RuntimeError):
    pass


def enabled() -> bool:
    return bool(settings.supabase_url and settings.supabase_service_role_key)


def _headers() -> dict[str, str]:
    return {
        "apikey": settings.supabase_service_role_key,
        "Authorization": f"Bearer {settings.supabase_service_role_key}",
        "Content-Type": "application/json",
    }


def save_latest(payload: dict[str, Any], workbook: bytes) -> None:
    if not enabled():
        return
    body = {"id": 1, "payload": payload, "workbook_base64": base64.b64encode(workbook).decode("ascii")}
    try:
        response = httpx.post(
            f"{settings.supabase_url}/rest/v1/order_book_latest?on_conflict=id",
            headers={**_headers(), "Prefer": "resolution=merge-duplicates"},
            json=body,
            timeout=60,
        )
        response.raise_for_status()
    except httpx.HTTPError as exc:
        raise StorageError("Unable to save the latest workbook to permanent storage.") from exc


def load_latest() -> tuple[dict[str, Any], bytes] | None:
    if not enabled():
        return None
    try:
        response = httpx.get(
            f"{settings.supabase_url}/rest/v1/order_book_latest?id=eq.1&select=payload,workbook_base64",
            headers=_headers(),
            timeout=60,
        )
        response.raise_for_status()
        rows = response.json()
        if not rows:
            return None
        return rows[0]["payload"], base64.b64decode(rows[0]["workbook_base64"])
    except (httpx.HTTPError, KeyError, ValueError) as exc:
        raise StorageError("Unable to load the latest workbook from permanent storage.") from exc
