"""Görev talebi veritabanı işlemleri."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import UTC, datetime
from typing import Any, Dict, Iterable, List, Optional

from .form_service import get_connection


class TaskRequestError(Exception):
    """Görev talebi işlemlerinde oluşan hata."""


VALID_URGENCY = {"normal", "urgent", "very_urgent"}
VALID_STATUS = {"pending", "in_progress", "converted", "rejected"}


URGENCY_LABELS = {
    "normal": "Normal",
    "urgent": "Acil",
    "very_urgent": "Çok Acil",
}

URGENCY_BADGE_CLASS = {
    "normal": "urgency-normal",
    "urgent": "urgency-urgent",
    "very_urgent": "urgency-very-urgent",
}

STATUS_LABELS = {
    "pending": "Beklemede",
    "in_progress": "İncelemede",
    "converted": "Göreve Dönüştürüldü",
    "rejected": "Reddedildi",
}

STATUS_BADGE_CLASS = {
    "pending": "status-pending",
    "in_progress": "status-progress",
    "converted": "status-converted",
    "rejected": "status-rejected",
}


@dataclass
class TaskRequest:
    id: int
    customer_name: str
    customer_phone: Optional[str]
    customer_email: Optional[str]
    customer_address: Optional[str]
    request_description: str
    requirements: Optional[str]
    urgency: str
    requested_by_user_id: int
    status: str
    notes: Optional[str]
    assigned_to_user_id: Optional[int]
    converted_form_no: Optional[str]
    created_at: str
    updated_at: str
    requested_by_name: Optional[str]
    assigned_to_name: Optional[str]

    @property
    def display_id(self) -> str:
        return f"#{self.id:03d}"

    @property
    def urgency_label(self) -> str:
        return URGENCY_LABELS.get(self.urgency, self.urgency.title())

    @property
    def urgency_class(self) -> str:
        return URGENCY_BADGE_CLASS.get(self.urgency, "urgency-normal")

    @property
    def status_label(self) -> str:
        return STATUS_LABELS.get(self.status, self.status.title())

    @property
    def status_class(self) -> str:
        return STATUS_BADGE_CLASS.get(self.status, "status-pending")


def _now_str() -> str:
    return datetime.now(UTC).strftime("%Y-%m-%d %H:%M:%S")


def _row_to_request(row) -> TaskRequest:
    return TaskRequest(
        id=row["id"],
        customer_name=row["customer_name"],
        customer_phone=row["customer_phone"],
        customer_email=row["customer_email"],
        customer_address=row["customer_address"],
        request_description=row["request_description"],
        requirements=row["requirements"],
        urgency=row["urgency"],
        requested_by_user_id=row["requested_by_user_id"],
        status=row["status"],
        notes=row["notes"],
        assigned_to_user_id=row["assigned_to_user_id"],
        converted_form_no=row["converted_form_no"],
        created_at=row["created_at"],
        updated_at=row["updated_at"],
        requested_by_name=row["requested_by_name"],
        assigned_to_name=row["assigned_to_name"],
    )


def _format_summary(text: str, limit: int = 50) -> str:
    cleaned = (text or "").strip()
    if len(cleaned) <= limit:
        return cleaned
    return cleaned[: limit - 1].rstrip() + "…"


def _format_datetime(value: str) -> str:
    if not value:
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f"):
        try:
            parsed = datetime.strptime(value, fmt)
        except ValueError:
            continue
        else:
            return parsed.strftime("%d.%m.%Y %H:%M")
    return value


def create_task_request(
    *,
    customer_name: str,
    customer_phone: Optional[str],
    customer_email: Optional[str],
    customer_address: Optional[str],
    request_description: str,
    requirements: Optional[str],
    urgency: str,
    requested_by_user_id: int,
    status: str = "pending",
    notes: Optional[str] = None,
    assigned_to_user_id: Optional[int] = None,
    converted_form_no: Optional[str] = None,
    base_path: str = ".",
) -> Dict[str, Any]:
    urgency = (urgency or "normal").strip().lower()
    if urgency not in VALID_URGENCY:
        urgency = "normal"

    status = (status or "pending").strip().lower()
    if status not in VALID_STATUS:
        status = "pending"

    now = _now_str()

    with get_connection(base_path) as connection:
        cursor = connection.execute(
            """
            INSERT INTO task_requests (
                customer_name,
                customer_phone,
                customer_email,
                customer_address,
                request_description,
                requirements,
                urgency,
                requested_by_user_id,
                status,
                notes,
                assigned_to_user_id,
                converted_form_no,
                created_at,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                customer_name,
                customer_phone,
                customer_email,
                customer_address,
                request_description,
                requirements,
                urgency,
                requested_by_user_id,
                status,
                notes,
                assigned_to_user_id,
                converted_form_no,
                now,
                now,
            ),
        )
        request_id = cursor.lastrowid
        connection.commit()

    created = get_task_request(request_id, base_path=base_path)
    if created is None:
        raise TaskRequestError("Görev talebi oluşturulamadı.")
    return created


def list_task_requests(*, status: Optional[str] = None, base_path: str = ".") -> List[Dict[str, Any]]:
    status_filter: Iterable[Any] = ()
    where_clause = ""
    if status:
        normalized = status.strip().lower()
        if normalized in VALID_STATUS:
            status_filter = (normalized,)
            where_clause = " WHERE tr.status = ?"

    query = (
        """
        SELECT tr.*, req.full_name AS requested_by_name, ass.full_name AS assigned_to_name
        FROM task_requests AS tr
        LEFT JOIN users AS req ON req.id = tr.requested_by_user_id
        LEFT JOIN users AS ass ON ass.id = tr.assigned_to_user_id
        """
        + where_clause
        + " ORDER BY tr.created_at DESC"
    )

    with get_connection(base_path) as connection:
        rows = connection.execute(query, tuple(status_filter)).fetchall()

    requests: List[Dict[str, Any]] = []
    for row in rows:
        request = _row_to_request(row)
        requests.append(
            {
                "id": request.id,
                "display_id": request.display_id,
                "customer_name": request.customer_name,
                "customer_phone": request.customer_phone,
                "customer_email": request.customer_email,
                "customer_address": request.customer_address,
                "request_description": request.request_description,
                "request_summary": _format_summary(request.request_description),
                "requirements": request.requirements,
                "urgency": request.urgency,
                "urgency_label": request.urgency_label,
                "urgency_class": request.urgency_class,
                "status": request.status,
                "status_label": request.status_label,
                "status_class": request.status_class,
                "requested_by_user_id": request.requested_by_user_id,
                "requested_by_name": request.requested_by_name,
                "assigned_to_user_id": request.assigned_to_user_id,
                "assigned_to_name": request.assigned_to_name,
                "converted_form_no": request.converted_form_no,
                "notes": request.notes or "",
                "created_at": request.created_at,
                "created_display": _format_datetime(request.created_at),
                "updated_at": request.updated_at,
            }
        )
    return requests


def get_task_request(request_id: int, *, base_path: str = ".") -> Optional[Dict[str, Any]]:
    query = """
        SELECT tr.*, req.full_name AS requested_by_name, ass.full_name AS assigned_to_name
        FROM task_requests AS tr
        LEFT JOIN users AS req ON req.id = tr.requested_by_user_id
        LEFT JOIN users AS ass ON ass.id = tr.assigned_to_user_id
        WHERE tr.id = ?
    """
    with get_connection(base_path) as connection:
        row = connection.execute(query, (request_id,)).fetchone()
    if row is None:
        return None
    request = _row_to_request(row)
    return {
        "id": request.id,
        "display_id": request.display_id,
        "customer_name": request.customer_name,
        "customer_phone": request.customer_phone,
        "customer_email": request.customer_email,
        "customer_address": request.customer_address,
        "request_description": request.request_description,
        "requirements": request.requirements,
        "urgency": request.urgency,
        "urgency_label": request.urgency_label,
        "urgency_class": request.urgency_class,
        "status": request.status,
        "status_label": request.status_label,
        "status_class": request.status_class,
        "requested_by_user_id": request.requested_by_user_id,
        "requested_by_name": request.requested_by_name,
        "assigned_to_user_id": request.assigned_to_user_id,
        "assigned_to_name": request.assigned_to_name,
        "converted_form_no": request.converted_form_no,
        "notes": request.notes or "",
        "created_at": request.created_at,
        "created_display": _format_datetime(request.created_at),
        "updated_at": request.updated_at,
    }


def update_task_request_status(
    request_id: int,
    *,
    status: str,
    base_path: str = ".",
) -> Dict[str, Any]:
    normalized = (status or "").strip().lower()
    if normalized not in VALID_STATUS:
        raise TaskRequestError("Geçersiz durum seçimi.")

    now = _now_str()

    with get_connection(base_path) as connection:
        cursor = connection.execute(
            "UPDATE task_requests SET status = ?, updated_at = ? WHERE id = ?",
            (normalized, now, request_id),
        )
        if cursor.rowcount == 0:
            raise TaskRequestError("Talep bulunamadı.")
        connection.commit()

    updated = get_task_request(request_id, base_path=base_path)
    if updated is None:
        raise TaskRequestError("Talep güncellenemedi.")
    return updated


def update_task_request_notes(
    request_id: int,
    *,
    notes: Optional[str],
    base_path: str = ".",
) -> Dict[str, Any]:
    cleaned = (notes or "").strip()
    now = _now_str()

    with get_connection(base_path) as connection:
        cursor = connection.execute(
            "UPDATE task_requests SET notes = ?, updated_at = ? WHERE id = ?",
            (cleaned or None, now, request_id),
        )
        if cursor.rowcount == 0:
            raise TaskRequestError("Talep bulunamadı.")
        connection.commit()

    updated = get_task_request(request_id, base_path=base_path)
    if updated is None:
        raise TaskRequestError("Talep güncellenemedi.")
    return updated


def mark_converted(
    request_id: int,
    *,
    form_no: str,
    base_path: str = ".",
) -> Dict[str, Any]:
    if not form_no:
        raise TaskRequestError("Form numarası zorunludur.")

    now = _now_str()

    with get_connection(base_path) as connection:
        cursor = connection.execute(
            """
            UPDATE task_requests
            SET status = 'converted', converted_form_no = ?, updated_at = ?
            WHERE id = ?
            """,
            (form_no, now, request_id),
        )
        if cursor.rowcount == 0:
            raise TaskRequestError("Talep bulunamadı.")
        connection.commit()

    updated = get_task_request(request_id, base_path=base_path)
    if updated is None:
        raise TaskRequestError("Talep güncellenemedi.")
    return updated


def get_pending_requests_count(*, base_path: str = ".") -> int:
    with get_connection(base_path) as connection:
        row = connection.execute(
            "SELECT COUNT(*) AS total FROM task_requests WHERE status = 'pending'"
        ).fetchone()
    if row is None:
        return 0
    return int(row["total"] or 0)
