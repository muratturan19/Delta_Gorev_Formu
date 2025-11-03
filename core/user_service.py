"""Kullanıcı veritabanı işlemleri."""
from __future__ import annotations

import os

from dataclasses import dataclass
from typing import Any, Iterable, List, Optional, Sequence

from werkzeug.security import check_password_hash, generate_password_hash

from .form_service import get_connection

DEFAULT_ASSIGNER_PASSWORD = os.environ.get("DEFAULT_ASSIGNER_PASSWORD", "Gorev123!")


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
    return list_users_by_roles((role,), base_path=base_path)


def list_users_by_roles(roles: Sequence[str], *, base_path: str = ".") -> List[User]:
    normalized_list: List[str] = []
    for role in roles:
        normalized = role.strip().lower() if role else ""
        if not normalized or normalized in normalized_list:
            continue
        normalized_list.append(normalized)
    if not normalized_list:
        return []
    placeholders = ",".join("?" for _ in normalized_list)
    query = (
        "SELECT * FROM users WHERE role IN ("
        + placeholders
        + ") AND is_active = 1 ORDER BY full_name COLLATE NOCASE"
    )
    with get_connection(base_path) as connection:
        rows = connection.execute(query, tuple(normalized_list)).fetchall()
    return [_row_to_user(row) for row in rows]


def get_user(user_id: int, *, base_path: str = ".") -> Optional[User]:
    with get_connection(base_path) as connection:
        row = connection.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    if row is None:
        return None
    return _row_to_user(row)


def get_user_by_name(full_name: str, *, base_path: str = ".") -> Optional[User]:
    """Ekip üyesi rolündeki kullanıcıyı tam ismine göre döndür."""

    normalized = (full_name or "").strip()
    if not normalized:
        return None

    with get_connection(base_path) as connection:
        row = connection.execute(
            """
            SELECT *
            FROM users
            WHERE full_name = ? AND role = 'calisan' AND is_active = 1
            LIMIT 1
            """,
            (normalized,),
        ).fetchone()

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
        # Legacy hazırlayan listesi (ilk üç isim)
        {
            "full_name": "Ali Yılmaz",
            "email": None,
            "phone": None,
            "password": DEFAULT_ASSIGNER_PASSWORD,
            "role": "atayan",
        },
        {
            "full_name": "Ayşe Demir",
            "email": None,
            "phone": None,
            "password": DEFAULT_ASSIGNER_PASSWORD,
            "role": "atayan",
        },
        {
            "full_name": "Mehmet Korkmaz",
            "email": None,
            "phone": None,
            "password": DEFAULT_ASSIGNER_PASSWORD,
            "role": "atayan",
        },
        # Legacy personel listesi
        {"full_name": "Ahmet Yılmaz", "email": None, "phone": None, "password": None, "role": "calisan"},
        {"full_name": "Mehmet Demir", "email": None, "phone": None, "password": None, "role": "calisan"},
        {"full_name": "Ali Kaya", "email": None, "phone": None, "password": None, "role": "calisan"},
        {"full_name": "Veli Çelik", "email": None, "phone": None, "password": None, "role": "calisan"},
        {"full_name": "Hasan Şahin", "email": None, "phone": None, "password": None, "role": "calisan"},
        {"full_name": "Hüseyin Aydın", "email": None, "phone": None, "password": None, "role": "calisan"},
        {"full_name": "İbrahim Özdemir", "email": None, "phone": None, "password": None, "role": "calisan"},
        {"full_name": "Mustafa Arslan", "email": None, "phone": None, "password": None, "role": "calisan"},
        {"full_name": "Emre Doğan", "email": None, "phone": None, "password": None, "role": "calisan"},
        {"full_name": "Burak Yıldız", "email": None, "phone": None, "password": None, "role": "calisan"},
        # Önceki varsayılan ekip üyeleri
        {
            "full_name": "Mehmet Ekip",
            "email": "calisan1@deltaproje.com",
            "phone": "5551234567",
            "password": None,
            "role": "calisan",
        },
        {
            "full_name": "Ayşe Ekip",
            "email": "calisan2@deltaproje.com",
            "phone": "5559876543",
            "password": None,
            "role": "calisan",
        },
        {
            "full_name": "Ali Ekip",
            "email": "calisan3@deltaproje.com",
            "phone": "5555555555",
            "password": None,
            "role": "calisan",
        },
    ]

    with get_connection(base_path) as connection:
        row = connection.execute("SELECT COUNT(*) AS total FROM users").fetchone()
        total_users = (row["total"] if row is not None else 0) or 0

    if total_users > 0:
        return

    for payload in defaults:
        try:
            create_user(
                full_name=payload["full_name"],
                email=payload["email"],
                phone=payload["phone"],
                password=payload["password"],
                role=payload["role"],
                base_path=base_path,
            )
        except UserServiceError:
            continue


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
