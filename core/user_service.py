"""Kullanıcı veritabanı işlemleri."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable, List, Optional

from werkzeug.security import check_password_hash, generate_password_hash

from .form_service import get_connection


class UserServiceError(Exception):
    """Kullanıcı işlemleri için genel hata."""


@dataclass
class User:
    id: int
    full_name: str
    email: Optional[str]
    phone: Optional[str]
    role: str
    is_active: bool

    @property
    def requires_password(self) -> bool:
        return self.role in {"admin", "atayan"}


def _row_to_user(row) -> User:
    return User(
        id=row["id"],
        full_name=row["full_name"],
        email=row["email"],
        phone=row["phone"],
        role=row["role"],
        is_active=bool(row["is_active"]),
    )


def list_users(*, base_path: str = ".", include_inactive: bool = False) -> List[User]:
    query = "SELECT * FROM users"
    params: Iterable[Any] = ()
    if not include_inactive:
        query += " WHERE is_active = 1"

    query += " ORDER BY full_name COLLATE NOCASE"

    with get_connection(base_path) as connection:
        rows = connection.execute(query, tuple(params)).fetchall()

    return [_row_to_user(row) for row in rows]


def list_users_by_role(role: str, *, base_path: str = ".") -> List[User]:
    with get_connection(base_path) as connection:
        rows = connection.execute(
            "SELECT * FROM users WHERE role = ? AND is_active = 1 ORDER BY full_name COLLATE NOCASE",
            (role,),
        ).fetchall()
    return [_row_to_user(row) for row in rows]


def get_user(user_id: int, *, base_path: str = ".") -> Optional[User]:
    with get_connection(base_path) as connection:
        row = connection.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    if row is None:
        return None
    return _row_to_user(row)


def authenticate_user(user_id: int, password: str, *, base_path: str = ".") -> bool:
    with get_connection(base_path) as connection:
        row = connection.execute(
            "SELECT password_hash, role FROM users WHERE id = ? AND is_active = 1",
            (user_id,),
        ).fetchone()
    if row is None:
        return False

    password_hash = row["password_hash"] or ""
    role = row["role"] or ""
    if role not in {"admin", "atayan"}:
        return True
    if not password_hash:
        return False
    return check_password_hash(password_hash, password)


def create_user(
    *,
    full_name: str,
    email: Optional[str],
    phone: Optional[str],
    password: Optional[str],
    role: str,
    base_path: str = ".",
) -> User:
    full_name = (full_name or "").strip()
    if not full_name:
        raise UserServiceError("İsim alanı zorunludur.")

    role = (role or "").strip().lower()
    if role not in {"admin", "atayan", "calisan"}:
        raise UserServiceError("Geçersiz rol seçimi.")

    email = (email or "").strip() or None
    phone = (phone or "").strip() or None

    password_hash = None
    if role in {"admin", "atayan"}:
        if not password or len(password) < 8:
            raise UserServiceError("Şifre en az 8 karakter olmalıdır.")
        password_hash = generate_password_hash(password)

    with get_connection(base_path) as connection:
        cursor = connection.execute(
            """
            INSERT INTO users (full_name, email, phone, password_hash, role, is_active)
            VALUES (?, ?, ?, ?, ?, 1)
            """,
            (full_name, email, phone, password_hash, role),
        )
        user_id = cursor.lastrowid
        connection.commit()

    created = get_user(user_id, base_path=base_path)
    if created is None:
        raise UserServiceError("Kullanıcı oluşturulamadı.")
    return created


def delete_user(user_id: int, *, base_path: str = ".") -> None:
    with get_connection(base_path) as connection:
        connection.execute("DELETE FROM users WHERE id = ?", (user_id,))
        connection.commit()


def ensure_default_users(*, base_path: str = ".") -> None:
    defaults = [
        {
            "full_name": "Admin User",
            "email": "admin@deltaproje.com",
            "phone": None,
            "password": "Delta2025!",
            "role": "admin",
        },
        {
            "full_name": "Ahmet Yönetici",
            "email": "yonetici@deltaproje.com",
            "phone": None,
            "password": "Yonetici123!",
            "role": "atayan",
        },
        {
            "full_name": "Mehmet Çalışan",
            "email": "calisan1@deltaproje.com",
            "phone": "5551234567",
            "password": None,
            "role": "calisan",
        },
        {
            "full_name": "Ayşe Çalışan",
            "email": "calisan2@deltaproje.com",
            "phone": "5559876543",
            "password": None,
            "role": "calisan",
        },
        {
            "full_name": "Ali Çalışan",
            "email": "calisan3@deltaproje.com",
            "phone": "5555555555",
            "password": None,
            "role": "calisan",
        },
    ]

    with get_connection(base_path) as connection:
        existing = connection.execute(
            "SELECT full_name, email FROM users"
        ).fetchall()
        existing_pairs = {(row["full_name"], row["email"]) for row in existing}

    for payload in defaults:
        identifier = (payload["full_name"], payload["email"])
        if identifier in existing_pairs:
            continue
        create_user(
            full_name=payload["full_name"],
            email=payload["email"],
            phone=payload["phone"],
            password=payload["password"],
            role=payload["role"],
            base_path=base_path,
        )


def update_user_password(user_id: int, password: str, *, base_path: str = ".") -> None:
    if len(password) < 8:
        raise UserServiceError("Şifre en az 8 karakter olmalıdır.")
    password_hash = generate_password_hash(password)
    with get_connection(base_path) as connection:
        connection.execute(
            "UPDATE users SET password_hash = ? WHERE id = ?",
            (password_hash, user_id),
        )
        connection.commit()
