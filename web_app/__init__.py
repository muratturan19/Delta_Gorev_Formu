# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Set
from uuid import uuid4

import jwt as pyjwt
from flask import (Flask, flash, g, redirect, render_template, request,
                   send_file, send_from_directory, session, url_for)
from werkzeug.middleware.proxy_fix import ProxyFix
from werkzeug.utils import secure_filename

from core import form_service, task_request_service, user_service
from core.form_service import FormServiceError
from core.user_service import UserServiceError

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
DEV_MODE = os.environ.get("DEV_MODE", "1").strip().lower() in {"1", "true", "yes", "on"}
SECRET_KEY = os.environ.get("SECRET_KEY", "dev-secret-key")

BASE_PATH = Path(__file__).resolve().parents[1]
DATA_FOLDER = os.environ.get("DATA_FOLDER", "").strip()
DATA_FILE = Path(DATA_FOLDER) / "data.json" if DATA_FOLDER else BASE_PATH / "data.json"
DEFAULT_FORM_VALUES: Dict[str, Any] = {
    "dok_no": "F-001",
    "rev_no": "00 / 06.05.24",
    "avans": "",
    "taseron": "",
    "gorev_tanimi": "",
    "gorev_yeri": "",
    "gorev_il": "",
    "gorev_ilce": "",
    "gorev_firma": "",
    "yapilan_isler": "",
    "yola_cikis_tarih": "",
    "yola_cikis_saat": "",
    "donus_tarih": "",
    "donus_saat": "",
    "calisma_baslangic_tarih": "",
    "calisma_baslangic_saat": "",
    "calisma_bitis_tarih": "",
    "calisma_bitis_saat": "",
    "mola_suresi": "",
    "arac_plaka": "",
    "hazirlayan": "",
    "gorev_tarih": "",
    "durum": "YARIM",
    "gorev_ekleri": [],
    "harcama_bildirimleri": [],
}
DEFAULT_FORM_VALUES.update({f"personel_{index}": "" for index in range(1, 6)})
DEFAULT_FORM_VALUES.update(
    {
        "assigned_to_user_id": None,
        "assigned_by_user_id": None,
        "assigned_at": None,
    }
)

FIELD_LABELS: Dict[str, str] = {
    "dok_no": "DOK.NO",
    "rev_no": "REV.NO/TRH",
    "avans": "Avans Tutarı",
    "taseron": "Taşeron Şirket",
    "gorev_tanimi": "Görevin Tanımı",
    "gorev_yeri": "Görev Yeri",
    "gorev_il": "Görev İli",
    "gorev_ilce": "Görev İlçesi",
    "gorev_firma": "Görev/Firma",
    "yola_cikis_tarih": "Yola Çıkış Tarihi",
    "yola_cikis_saat": "Yola Çıkış Saati",
    "donus_tarih": "Dönüş Tarihi",
    "donus_saat": "Dönüş Saati",
    "calisma_baslangic_tarih": "Çalışma Başlangıç Tarihi",
    "calisma_baslangic_saat": "Çalışma Başlangıç Saati",
    "calisma_bitis_tarih": "Çalışma Bitiş Tarihi",
    "calisma_bitis_saat": "Çalışma Bitiş Saati",
    "mola_suresi": "Toplam Mola",
    "arac_plaka": "Araç Plaka No",
    "hazirlayan": "Hazırlayan",
    "gorev_tarih": "Görev Tarihi",
}

DEFAULT_DYNAMIC_DATA: Dict[str, List[str]] = {
    "taseron_options": [
        "Yok",
        "ABC İnşaat",
        "XYZ Teknik",
        "Marmara Mühendislik",
        "Anadolu Yapı",
    ],
    "arac_plaka_options": [
        "34 ABC 123",
        "34 XYZ 456",
        "06 DEF 789",
        "16 GHI 321",
        "35 JKL 654",
    ],
}

DEFAULT_FORM_DEFAULTS: Dict[str, str] = {
    "dok_no": "F-001",
    "rev_no": "00 / 06.05.24",
}

_upload_folder = os.environ.get("UPLOAD_FOLDER", "").strip()
UPLOAD_DIR = Path(_upload_folder) if _upload_folder else BASE_PATH / "uploads"

_storage: Dict[str, Any] | None = None
_dynamic_data: Dict[str, List[str]] | None = None
_form_defaults: Dict[str, str] | None = None
ADMIN_PASSWORD = os.environ.get("ADMIN_PANEL_PASSWORD") or os.environ.get("ADMIN_PASSWORD") or "delta-admin"

FORM_STEPS: List[Dict[str, str]] = [
    {"id": "form_bilgileri", "title": "Form Bilgileri", "template": "steps/form_bilgileri.html"},
    {"id": "hazirlayan", "title": "Hazırlayan", "template": "steps/hazirlayan.html"},
    {"id": "gorevli_personel", "title": "Görevli Personel", "template": "steps/gorevli_personel.html"},
    {"id": "finans_arac", "title": "Avans ve Araç Bilgileri", "template": "steps/finans_arac.html"},
    {"id": "gorev_detay", "title": "Görev Tanımı ve Yeri", "template": "steps/gorev_detay.html"},
    {"id": "gorev_bilgileri", "title": "Görev Bilgileri", "template": "steps/gorev_bilgileri.html"},
]


def normalize_options(values) -> List[str]:
    if not isinstance(values, list):
        return []
    seen = set()
    normalized: List[str] = []
    for item in values:
        if not isinstance(item, str):
            continue
        cleaned = item.strip()
        if not cleaned or cleaned in seen:
            continue
        seen.add(cleaned)
        normalized.append(cleaned)
    return normalized


def load_storage() -> Dict[str, Any]:
    if not DATA_FILE.exists():
        return {}
    try:
        with DATA_FILE.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(payload, dict):
        return {}
    return payload


def save_storage(payload: Dict[str, Any]) -> None:
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    with DATA_FILE.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)


def get_storage() -> Dict[str, Any]:
    global _storage
    if _storage is None:
        _storage = load_storage()
    return _storage


def set_storage(payload: Dict[str, Any]) -> None:
    global _storage
    _storage = payload
    save_storage(payload)


def load_dynamic_data() -> Dict[str, List[str]]:
    storage = get_storage()
    data = {key: list(value) for key, value in DEFAULT_DYNAMIC_DATA.items()}
    for key, default_values in DEFAULT_DYNAMIC_DATA.items():
        values = normalize_options(storage.get(key, default_values))
        data[key] = values or list(default_values)
    return data


def save_dynamic_data(data: Dict[str, List[str]]) -> None:
    storage = get_storage().copy()
    storage.update(data)
    set_storage(storage)


def migrate_legacy_user_lists(
    *, base_path: str, assigners: Optional[List[str]] = None, employees: Optional[List[str]] = None
) -> None:
    storage = get_storage()

    legacy_assigners = assigners if assigners is not None else storage.get("hazirlayan_options", [])
    legacy_employees = employees if employees is not None else storage.get("personel_options", [])

    assigner_names: Set[str] = {
        user.full_name.strip().lower()
        for user in user_service.list_users_by_roles(("admin", "atayan"), base_path=base_path)
    }
    employee_names: Set[str] = {
        user.full_name.strip().lower()
        for user in user_service.list_users_by_role("calisan", base_path=base_path)
    }

    created_assigners = 0
    if isinstance(legacy_assigners, list):
        for raw_name in legacy_assigners:
            if created_assigners >= 3:
                break
            if not isinstance(raw_name, str):
                continue
            cleaned = raw_name.strip()
            if not cleaned or cleaned.lower() in assigner_names:
                continue
            try:
                user_service.create_user(
                    full_name=cleaned,
                    email=None,
                    phone=None,
                    password=user_service.DEFAULT_ASSIGNER_PASSWORD,
                    role="atayan",
                    base_path=base_path,
                )
            except UserServiceError:
                continue
            else:
                assigner_names.add(cleaned.lower())
                created_assigners += 1

    if isinstance(legacy_employees, list):
        for raw_name in legacy_employees:
            if not isinstance(raw_name, str):
                continue
            cleaned = raw_name.strip()
            if not cleaned or cleaned.lower() in employee_names:
                continue
            try:
                user_service.create_user(
                    full_name=cleaned,
                    email=None,
                    phone=None,
                    password=None,
                    role="calisan",
                    base_path=base_path,
                )
            except UserServiceError:
                continue
            else:
                employee_names.add(cleaned.lower())

    if assigners is None and employees is None:
        if "hazirlayan_options" in storage or "personel_options" in storage:
            cleaned_storage = storage.copy()
            cleaned_storage.pop("hazirlayan_options", None)
            cleaned_storage.pop("personel_options", None)
            set_storage(cleaned_storage)
            global _dynamic_data
            _dynamic_data = None


def get_dynamic_data() -> Dict[str, List[str]]:
    global _dynamic_data
    if _dynamic_data is None:
        _dynamic_data = load_dynamic_data()
    return _dynamic_data


def set_dynamic_data(data: Dict[str, List[str]]) -> None:
    global _dynamic_data
    normalized: Dict[str, List[str]] = {}
    for key, default_values in DEFAULT_DYNAMIC_DATA.items():
        normalized[key] = normalize_options(data.get(key, default_values)) or list(default_values)
    _dynamic_data = normalized
    save_dynamic_data(normalized)


def load_form_defaults() -> Dict[str, str]:
    storage = get_storage()
    stored = storage.get("form_defaults", {})
    defaults: Dict[str, str] = {}
    for key, fallback in DEFAULT_FORM_DEFAULTS.items():
        value = stored.get(key)
        if isinstance(value, str) and value.strip():
            defaults[key] = value.strip()
        else:
            defaults[key] = fallback
    return defaults


def save_form_defaults(defaults: Dict[str, str]) -> None:
    storage = get_storage().copy()
    storage["form_defaults"] = defaults
    set_storage(storage)


def get_form_defaults() -> Dict[str, str]:
    global _form_defaults
    if _form_defaults is None:
        _form_defaults = load_form_defaults()
    return _form_defaults


def set_form_defaults(defaults: Dict[str, str]) -> None:
    global _form_defaults
    cleaned: Dict[str, str] = {}
    for key, fallback in DEFAULT_FORM_DEFAULTS.items():
        value = defaults.get(key, "")
        if not isinstance(value, str):
            value = str(value or "")
        value = value.strip()
        cleaned[key] = value or fallback
    _form_defaults = cleaned
    save_form_defaults(cleaned)


def create_app() -> Flask:
    app = Flask(__name__)
    app.secret_key = os.environ.get("FLASK_SECRET_KEY", SECRET_KEY)

    # ProxyFix: trust X-Forwarded-* headers from reverse proxy (Nginx)
    app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_prefix=1)

    if DEV_MODE:
        user_service.ensure_default_users(base_path=str(BASE_PATH))
        migrate_legacy_user_lists(base_path=str(BASE_PATH))

    register_routes(app)
    return app


def _auto_provision_user(payload: Dict[str, Any]) -> Dict[str, Any]:
    """JWT payload'ından kullanıcıyı oluştur veya güncelle, dict olarak döndür."""
    portal_user_id = payload.get("user_id")
    username = payload.get("username", "Bilinmeyen Kullanıcı")
    role = payload.get("role", "calisan")
    email = payload.get("email", "")

    role_map = {"admin": "admin", "atayan": "atayan", "calisan": "calisan"}
    local_role = role_map.get(role, "calisan")

    existing = user_service.get_user_by_portal_id(portal_user_id, base_path=str(BASE_PATH))
    if existing:
        if existing.role != local_role:
            user_service.update_user_role(existing.id, local_role, base_path=str(BASE_PATH))
        return {
            "id": existing.id,
            "full_name": existing.full_name,
            "email": existing.email or email,
            "role": local_role,
        }

    new_user = user_service.create_user(
        full_name=username,
        email=email or None,
        phone=None,
        password=None,
        role=local_role,
        portal_user_id=portal_user_id,
        base_path=str(BASE_PATH),
    )
    return {
        "id": new_user.id,
        "full_name": new_user.full_name,
        "email": new_user.email or email,
        "role": new_user.role,
    }


def register_routes(app: Flask) -> None:
    total_steps = len(FORM_STEPS)

    def get_current_user() -> Dict[str, Any] | None:
        # Production: JWT middleware sets g.user
        if hasattr(g, "user") and g.user:
            return g.user
        # Dev mode: session-based auth
        user_data = session.get("user")
        if not isinstance(user_data, dict):
            return None
        return user_data

    def current_role() -> str | None:
        user_data = get_current_user()
        if not user_data:
            return None
        return user_data.get("role")

    def has_role(*roles: str) -> bool:
        role = current_role()
        return role in roles if role else False

    def require_roles(*roles: str):
        response = require_login()
        if response is not None:
            return response
        if not has_role(*roles):
            flash("Bu işlemi gerçekleştirmek için yetkiniz yok.", "error")
            return redirect(url_for("index"))
        return None

    def require_login():
        if get_current_user() is None:
            if DEV_MODE:
                flash("Lütfen önce kendinizi seçin.", "warning")
                return redirect(url_for("index"))
            return "Oturum bulunamadı. Portal üzerinden giriş yapın.", 401
        return None

    def set_session_user(user_obj) -> None:
        session["user"] = {
            "id": user_obj.id,
            "full_name": user_obj.full_name,
            "email": user_obj.email,
            "role": user_obj.role,
        }

    def clear_session_user() -> None:
        session.pop("user", None)

    def normalize_expense_entries(form_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        entries = form_data.get("harcama_bildirimleri", [])
        normalized_entries: List[Dict[str, Any]] = []
        if isinstance(entries, list):
            for item in entries:
                if not isinstance(item, dict):
                    continue
                description = (item.get("description") or "").strip()
                attachments: List[Dict[str, str]] = []
                raw_attachments = item.get("attachments", [])
                if isinstance(raw_attachments, list):
                    for attachment in raw_attachments:
                        if isinstance(attachment, dict) and attachment.get("filename"):
                            attachments.append(
                                {
                                    "filename": attachment["filename"],
                                    "original_name": attachment.get("original_name")
                                    or attachment["filename"],
                                }
                            )
                normalized_entries.append({"description": description, "attachments": attachments})
        form_data["harcama_bildirimleri"] = normalized_entries
        return normalized_entries

    def normalize_attachments(form_data: Dict[str, Any]) -> List[Dict[str, str]]:
        attachments = form_data.get("gorev_ekleri", [])
        normalized: List[Dict[str, str]] = []
        if isinstance(attachments, list):
            for item in attachments:
                if isinstance(item, dict) and "filename" in item:
                    normalized.append(
                        {
                            "filename": item["filename"],
                            "original_name": item.get("original_name") or item["filename"],
                        }
                    )
        form_data["gorev_ekleri"] = normalized
        normalize_expense_entries(form_data)
        return normalized

    def save_uploaded_files(form_no: str, files, field_name: str) -> List[Dict[str, str]]:
        if not files:
            return []
        try:
            file_items = files.getlist(field_name)
        except AttributeError:
            file_items = []
        uploaded: List[Dict[str, str]] = []
        for file in file_items:
            if not file or not file.filename:
                continue
            filename = secure_filename(file.filename)
            if not filename:
                continue
            unique_name = f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{uuid4().hex[:8]}_{filename}"
            target_dir = UPLOAD_DIR / form_no
            target_dir.mkdir(parents=True, exist_ok=True)
            target_path = target_dir / unique_name
            file.save(target_path)
            uploaded.append({"filename": unique_name, "original_name": file.filename})
        return uploaded

    def delete_attachment_file(form_no: str, filename: str) -> bool:
        if not filename:
            return False
        target_path = UPLOAD_DIR / form_no / filename
        if not target_path.exists():
            return False
        try:
            target_path.unlink()
        except OSError:
            return False
        return True

    def ensure_admin_access():
        response = require_login()
        if response is not None:
            return response
        if not has_role("admin"):
            flash("Bu bölüme yalnızca admin kullanıcılar erişebilir.", "error")
            return redirect(url_for("index"))
        return None

    # -- JWT authentication middleware (production) -------------------------
    @app.before_request
    def jwt_auth():
        if DEV_MODE:
            return  # Dev mode: session-based auth, skip JWT

        # Allow static file requests without auth
        if request.endpoint == "static":
            return

        token = request.cookies.get("delta_token")
        if not token:
            g.user = None
            return

        try:
            payload = pyjwt.decode(token, SECRET_KEY, algorithms=["HS256"])
        except pyjwt.ExpiredSignatureError:
            g.user = None
            return
        except pyjwt.InvalidTokenError:
            g.user = None
            return

        g.user = _auto_provision_user(payload)

    @app.template_filter("to_html_date")
    def to_html_date(value: str) -> str:
        value = (value or "").strip()
        if not value:
            return ""
        for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
            try:
                parsed = datetime.strptime(value, fmt)
            except ValueError:
                continue
            else:
                return parsed.strftime("%Y-%m-%d")
        return value

    @app.context_processor
    def inject_user():  # pragma: no cover - template helper
        return {"current_user": get_current_user()}

    @app.context_processor
    def inject_globals():  # pragma: no cover - template helper
        dynamic_data = get_dynamic_data()
        return {
            "FORM_STEPS": FORM_STEPS,
            "total_steps": total_steps,
            "taseron_options": dynamic_data["taseron_options"],
            "arac_plaka_options": dynamic_data["arac_plaka_options"],
            "is_admin": has_role("admin"),
        }

    @app.context_processor
    def inject_pending_requests():  # pragma: no cover - template helper
        user = get_current_user()
        if user and user.get("role") in ["admin", "atayan"]:
            pending_count = task_request_service.get_pending_requests_count(
                base_path=str(BASE_PATH)
            )
            return {"pending_requests_count": pending_count}
        return {"pending_requests_count": 0}

    def get_locked_forms() -> List[str]:
        return list(session.get("locked_forms", []))

    def set_locked_forms(values: List[str]) -> None:
        session["locked_forms"] = values

    def is_form_locked(form_no: str) -> bool:
        return form_no in set(get_locked_forms())

    def lock_form(form_no: str) -> None:
        locked = set(get_locked_forms())
        locked.add(form_no)
        set_locked_forms(sorted(locked))

    def unlock_form(form_no: str) -> None:
        locked = set(get_locked_forms())
        if form_no in locked:
            locked.remove(form_no)
            set_locked_forms(sorted(locked))

    def user_is_form_personnel(user: Dict[str, Any] | None, form_data: Dict[str, Any]) -> bool:
        if not user or user.get("role") != "calisan":
            return False

        user_name = (user.get("full_name") or "").strip()
        if not user_name:
            return False

        normalized_user = user_name.casefold()
        for field in form_service.PERSONEL_FIELDS:
            person_name = (form_data.get(field) or "").strip()
            if person_name and person_name.casefold() == normalized_user:
                return True
        return False

    @app.route("/")
    def index():
        current = get_current_user()
        all_users = user_service.list_users(base_path=str(BASE_PATH))
        pending_login = session.get("pending_login")
        show_password_modal = session.pop("show_password_modal", False)

        filters = {
            "personel": request.args.get("personel", "").strip(),
            "gorev_yeri": request.args.get("gorev_yeri", "").strip(),
            "start_date": request.args.get("start_date", "").strip(),
            "end_date": request.args.get("end_date", "").strip(),
        }
        search_triggered = request.args.get("performed_search", "").strip() == "1"

        form_numbers: List[str] = []
        search_results: List[Dict[str, Any]] = []
        performed_search = False
        assigned_forms: List[Dict[str, Any]] = []
        personnel_options: List[str] = []
        location_options: List[str] = []

        if current and current.get("role") == "calisan":
            filters = {key: "" for key in filters}
            assigned_forms = form_service.list_forms_for_assignee(
                current.get("id"),
                base_path=str(BASE_PATH),
                personnel_name=current.get("full_name"),
            )
            form_numbers = [item["form_no"] for item in assigned_forms]
        else:
            form_numbers = form_service.list_form_numbers(base_path=str(BASE_PATH))
            personnel_options = form_service.list_distinct_personnel(base_path=str(BASE_PATH))
            location_options = form_service.list_distinct_locations(base_path=str(BASE_PATH))

            current_person = filters["personel"]
            if current_person and not any(
                option.casefold() == current_person.casefold() for option in personnel_options
            ):
                personnel_options.append(current_person)
                personnel_options.sort(key=lambda item: item.casefold())

            current_location = filters["gorev_yeri"]
            if current_location and not any(
                option.casefold() == current_location.casefold() for option in location_options
            ):
                location_options.append(current_location)
                location_options.sort(key=lambda item: item.casefold())

            performed_search = search_triggered or any(filters.values())
            if performed_search:
                search_results = form_service.search_forms(
                    person=filters["personel"],
                    location=filters["gorev_yeri"],
                    start_date=filters["start_date"],
                    end_date=filters["end_date"],
                    base_path=str(BASE_PATH),
                )

        return render_template(
            "home.html",
            form_numbers=form_numbers,
            search_filters=filters,
            search_results=search_results,
            performed_search=performed_search,
            login_users=all_users,
            assigned_forms=assigned_forms,
            pending_login=pending_login,
            show_password_modal=show_password_modal,
            personel_options=personnel_options,
            location_options=location_options,
        )

    @app.post("/login/select")
    def login_select():
        user_id_raw = request.form.get("user_id", "").strip()
        if not user_id_raw.isdigit():
            flash("Lütfen bir kullanıcı seçin.", "warning")
            return redirect(url_for("index"))

        user_obj = user_service.get_user(int(user_id_raw), base_path=str(BASE_PATH))
        if user_obj is None or not user_obj.is_active:
            flash("Seçilen kullanıcı bulunamadı veya aktif değil.", "error")
            return redirect(url_for("index"))

        session.pop("pending_login", None)

        if user_obj.role == "calisan":
            set_session_user(user_obj)
            session.pop("show_password_modal", None)
            flash(f"Hoş geldiniz {user_obj.full_name}!", "success")
            return redirect(url_for("index"))

        session["pending_login"] = {"id": user_obj.id, "full_name": user_obj.full_name}
        session["show_password_modal"] = True
        flash(f"{user_obj.full_name} için şifre gerekli.", "info")
        return redirect(url_for("index"))

    @app.post("/login/password")
    def login_password():
        pending = session.get("pending_login")
        if not isinstance(pending, dict) or "id" not in pending:
            flash("Lütfen önce kullanıcı seçin.", "warning")
            return redirect(url_for("index"))

        password = request.form.get("password", "")
        user_obj = user_service.get_user(int(pending["id"]), base_path=str(BASE_PATH))
        if user_obj is None or not user_obj.requires_password:
            session.pop("pending_login", None)
            session.pop("show_password_modal", None)
            flash("Geçersiz giriş isteği.", "error")
            return redirect(url_for("index"))

        if user_service.authenticate_user(user_obj.id, password, base_path=str(BASE_PATH)):
            set_session_user(user_obj)
            session.pop("pending_login", None)
            session.pop("show_password_modal", None)
            flash(f"Hoş geldiniz {user_obj.full_name}!", "success")
            return redirect(url_for("index"))

        session["pending_login"] = {"id": user_obj.id, "full_name": user_obj.full_name}
        session["show_password_modal"] = True
        flash("Şifre hatalı. Lütfen tekrar deneyin.", "error")
        return redirect(url_for("index"))

    @app.get("/login/cancel")
    def login_cancel():
        session.pop("pending_login", None)
        session.pop("show_password_modal", None)
        flash("Giriş işlemi iptal edildi.", "info")
        return redirect(url_for("index"))

    @app.route("/task-request/new", methods=["GET", "POST"])
    def task_request_new():
        response = require_login()
        if response is not None:
            return response

        current = get_current_user()
        form_values = {
            "customer_name": "",
            "customer_phone": "",
            "customer_email": "",
            "customer_address": "",
            "request_description": "",
            "requirements": "",
            "urgency": "normal",
        }
        errors: Dict[str, str] = {}

        if request.method == "POST":
            form_values.update(
                {
                    "customer_name": request.form.get("customer_name", "").strip(),
                    "customer_phone": request.form.get("customer_phone", "").strip(),
                    "customer_email": request.form.get("customer_email", "").strip(),
                    "customer_address": request.form.get("customer_address", "").strip(),
                    "request_description": request.form.get("request_description", "").strip(),
                    "requirements": request.form.get("requirements", "").strip(),
                    "urgency": (request.form.get("urgency", "normal") or "normal").strip().lower(),
                }
            )

            if len(form_values["customer_name"]) < 2:
                errors["customer_name"] = "Lütfen müşteri adını girin (en az 2 karakter)."

            if form_values["customer_phone"]:
                import re

                phone_pattern = re.compile(r"^0\d{3}\s?\d{3}\s?\d{2}\s?\d{2}$")
                if not phone_pattern.match(form_values["customer_phone"]):
                    errors["customer_phone"] = "Telefon numarası 0xxx xxx xx xx formatında olmalıdır."

            if len(form_values["request_description"]) < 10:
                errors["request_description"] = "Talep detayı en az 10 karakter olmalıdır."

            if len(form_values["requirements"]) > 500:
                errors["requirements"] = "Gereklilikler en fazla 500 karakter olabilir."

            if form_values["urgency"] not in task_request_service.VALID_URGENCY:
                form_values["urgency"] = "normal"

            if not errors:
                try:
                    task_request_service.create_task_request(
                        customer_name=form_values["customer_name"],
                        customer_phone=form_values["customer_phone"] or None,
                        customer_email=form_values["customer_email"] or None,
                        customer_address=form_values["customer_address"] or None,
                        request_description=form_values["request_description"],
                        requirements=form_values["requirements"] or None,
                        urgency=form_values["urgency"],
                        requested_by_user_id=current.get("id"),
                        base_path=str(BASE_PATH),
                    )
                except task_request_service.TaskRequestError as exc:  # pragma: no cover - güvenlik
                    flash(str(exc), "error")
                else:
                    flash("Görev talebi oluşturuldu. İlgili ekip bilgilendirildi.", "success")
                    return redirect(url_for("index"))
            else:
                flash("Lütfen formu kontrol edin.", "error")

        urgency_options = [
            {"value": "normal", "label": "Normal"},
            {"value": "urgent", "label": "Acil"},
            {"value": "very_urgent", "label": "Çok Acil"},
        ]

        return render_template(
            "task_request_form.html",
            form_values=form_values,
            form_errors=errors,
            urgency_options=urgency_options,
        )

    @app.route("/reports")
    def reports():
        response = require_roles("admin", "atayan")
        if response is not None:
            return response

        selected_start = request.args.get("start_date", "").strip()
        selected_end = request.args.get("end_date", "").strip()

        today = datetime.utcnow().date()
        month_start = today.replace(day=1)
        if month_start.month == 12:
            next_month = month_start.replace(year=month_start.year + 1, month=1)
        else:
            next_month = month_start.replace(month=month_start.month + 1)
        month_end = next_month - timedelta(days=1)

        year_start = today.replace(month=1, day=1)
        year_end = year_start.replace(year=year_start.year + 1) - timedelta(days=1)

        if not selected_start and not selected_end:
            selected_start = month_start.isoformat()
            selected_end = month_end.isoformat()

        summary = form_service.get_reporting_summary(
            start_date=selected_start,
            end_date=selected_end,
            base_path=str(BASE_PATH),
        )

        quick_filters = {
            "month": {
                "label": "Bu Ay",
                "start": month_start.isoformat(),
                "end": month_end.isoformat(),
            },
            "year": {
                "label": "Bu Yıl",
                "start": year_start.isoformat(),
                "end": year_end.isoformat(),
            },
        }

        return render_template(
            "reports.html",
            report=summary,
            selected_start=selected_start,
            selected_end=selected_end,
            quick_filters=quick_filters,
        )

    @app.route("/task-requests")
    def task_requests_list():
        response = require_roles("admin", "atayan")
        if response is not None:
            return response

        status_filter = (request.args.get("status", "") or "").strip().lower()
        allowed_filters = {"pending", "converted", "rejected"}
        if status_filter not in allowed_filters:
            status_filter = ""

        requests_data = task_request_service.list_task_requests(
            status=status_filter or None,
            base_path=str(BASE_PATH),
        )

        for item in requests_data:
            item["update_url"] = url_for("task_request_update", request_id=item["id"])
            item["convert_url"] = url_for("new_form", from_request=item["id"])
            if item.get("converted_form_no"):
                item["converted_form_url"] = url_for(
                    "form_summary", form_no=item["converted_form_no"]
                )
            else:
                item["converted_form_url"] = ""

        status_options = [
            {"value": "", "label": "Tümü"},
            {"value": "pending", "label": "Bekleyen"},
            {"value": "converted", "label": "Göreve Dönüştürülenler"},
            {"value": "rejected", "label": "Reddedilenler"},
        ]

        return render_template(
            "task_requests_list.html",
            task_requests=requests_data,
            status_options=status_options,
            status_filter=status_filter,
        )

    @app.post("/task-requests/<int:request_id>/update")
    def task_request_update(request_id: int):
        response = require_roles("admin", "atayan")
        if response is not None:
            return response

        redirect_status = (request.form.get("redirect_status", "") or "").strip().lower()
        redirect_params = {}
        if redirect_status in {"pending", "converted", "rejected"}:
            redirect_params["status"] = redirect_status

        current_request = task_request_service.get_task_request(
            request_id, base_path=str(BASE_PATH)
        )
        if current_request is None:
            flash("Talep bulunamadı.", "error")
            return redirect(url_for("task_requests_list", **redirect_params))

        action = (request.form.get("action", "") or "").strip()
        if action == "update_status":
            status_value = (request.form.get("status_value", "") or "").strip().lower()
            allowed = False
            if status_value == "in_progress" and current_request["status"] == "pending":
                allowed = True
            elif status_value == "rejected" and current_request["status"] in {"pending", "in_progress"}:
                allowed = True

            if not allowed:
                flash("Bu durum değişikliği uygulanamaz.", "error")
                return redirect(url_for("task_requests_list", **redirect_params))

            try:
                updated = task_request_service.update_task_request_status(
                    request_id, status=status_value, base_path=str(BASE_PATH)
                )
            except task_request_service.TaskRequestError as exc:
                flash(str(exc), "error")
            else:
                message = (
                    "Talep incelemede olarak işaretlendi."
                    if updated["status"] == "in_progress"
                    else "Talep reddedildi."
                )
                flash(message, "success")
            return redirect(url_for("task_requests_list", **redirect_params))

        if action == "update_notes":
            notes = request.form.get("notes", "") or ""
            if len(notes) > 1000:
                flash("Notlar en fazla 1000 karakter olabilir.", "error")
                return redirect(url_for("task_requests_list", **redirect_params))
            try:
                task_request_service.update_task_request_notes(
                    request_id, notes=notes, base_path=str(BASE_PATH)
                )
            except task_request_service.TaskRequestError as exc:
                flash(str(exc), "error")
            else:
                flash("Notlar güncellendi.", "success")
            return redirect(url_for("task_requests_list", **redirect_params))

        flash("Geçersiz işlem.", "error")
        return redirect(url_for("task_requests_list", **redirect_params))

    @app.route("/gorevlerim")
    def my_tasks():
        response = require_roles("calisan")
        if response is not None:
            return response

        current = get_current_user()
        assignments = form_service.list_forms_for_assignee(
            current.get("id"),
            base_path=str(BASE_PATH),
            personnel_name=current.get("full_name"),
        )
        return render_template("my_tasks.html", assignments=assignments)

    @app.post("/form/load")
    def load_form_redirect():
        response = require_login()
        if response is not None:
            return response

        form_no = request.form.get("form_no_input", "").strip() or request.form.get("form_no_select", "").strip()
        if not form_no:
            flash("Form numarası seçin veya girin.", "warning")
            return redirect(url_for("index"))
        form_data = ensure_form_data(form_no)
        if form_data is None:
            return redirect(url_for("index"))

        current = get_current_user()
        is_employee = current and current.get("role") == "calisan"
        is_responsible = bool(
            is_employee and form_data.get("assigned_to_user_id") == current.get("id")
        )
        is_team_member = bool(is_employee and user_is_form_personnel(current, form_data))

        if is_employee and not (is_responsible or is_team_member):
            flash("Bu göreve erişiminiz yok.", "error")
            return redirect(url_for("index"))

        status = form_service.determine_form_status(form_data)
        form_data["durum"] = status.code
        store_form_in_session(form_no, form_data)

        if status.is_complete:
            lock_form(form_no)
            flash("Tamamlanmış form özet görünümünde açıldı.", "info")
            return redirect(url_for("form_summary", form_no=form_no))

        if is_employee and not is_responsible:
            flash(
                "Görev özetine yönlendirildiniz. Bu görevi yalnızca görev sorumlusu düzenleyebilir.",
                "info",
            )
            return redirect(url_for("form_summary", form_no=form_no))

        unlock_form(form_no)
        last_step = clamp_step(form_data.get("last_step", 0))
        form_data["last_step"] = last_step
        store_form_in_session(form_no, form_data)
        return redirect(url_for("form_wizard", form_no=form_no, step=last_step))

    @app.post("/form/<form_no>/assign")
    def assign_form_route(form_no: str):
        response = require_roles("admin", "atayan")
        if response is not None:
            return response

        form_data = ensure_form_data(form_no)
        if form_data is None:
            flash(f"Form {form_no} bulunamadı.", "error")
            return redirect(url_for("index"))

        assigned_id_raw = request.form.get("assigned_user_id", "").strip()
        assigned_id = int(assigned_id_raw) if assigned_id_raw.isdigit() else None
        if assigned_id is None:
            flash("Lütfen görevlendirmek için bir ekip üyesi seçin.", "warning")
            return redirect(url_for("form_wizard", form_no=form_no, step=4))

        employee = user_service.get_user(assigned_id, base_path=str(BASE_PATH))
        if employee is None or employee.role != "calisan":
            flash("Seçilen kullanıcı geçersiz.", "error")
            return redirect(url_for("form_wizard", form_no=form_no, step=4))

        current = get_current_user()

        try:
            _, status = form_service.save_partial_form(
                form_no, form_data, base_path=str(BASE_PATH)
            )
        except FormServiceError as exc:
            flash(str(exc), "error")
            return redirect(url_for("form_wizard", form_no=form_no, step=4))
        else:
            form_data["durum"] = status.code

        assigned_timestamp = form_service.assign_form(
            form_no,
            assigned_to_user_id=employee.id,
            assigned_by_user_id=current.get("id") if current else None,
            base_path=str(BASE_PATH),
        )

        form_data["assigned_to_user_id"] = employee.id
        form_data["assigned_by_user_id"] = current.get("id") if current else None
        form_data["assigned_at"] = assigned_timestamp
        store_form_in_session(form_no, form_data)

        flash(f"Form {employee.full_name} kullanıcısına atandı.", "success")
        return redirect(url_for("form_wizard", form_no=form_no, step=4))

    @app.route("/form/new")
    def new_form():
        response = require_roles("admin", "atayan")
        if response is not None:
            return response

        request_id_raw = (request.args.get("from_request", "") or "").strip()
        linked_request: Optional[Dict[str, Any]] = None
        if request_id_raw:
            if not request_id_raw.isdigit():
                flash("Geçersiz talep seçimi.", "error")
                return redirect(url_for("task_requests_list"))
            linked_request = task_request_service.get_task_request(
                int(request_id_raw), base_path=str(BASE_PATH)
            )
            if linked_request is None:
                flash("Talep bulunamadı.", "error")
                return redirect(url_for("task_requests_list"))
            if (
                linked_request.get("status") == "converted"
                and linked_request.get("converted_form_no")
            ):
                flash("Talep zaten göreve dönüştürülmüş.", "warning")
                return redirect(url_for("task_requests_list"))

        try:
            form_no = form_service.generate_form_number(base_path=str(BASE_PATH))
        except FormServiceError as exc:
            flash(str(exc), "error")
            return redirect(url_for("index"))

        if not form_no.startswith("F-"):
            form_no = f"F-{form_no}"

        form_defaults = get_form_defaults()
        form_data = {
            **DEFAULT_FORM_VALUES,
            "form_no": form_no,
            "tarih": datetime.now().strftime("%d.%m.%Y"),
            "dok_no": form_defaults["dok_no"],
            "rev_no": form_defaults["rev_no"],
        }
        form_data["gorev_ekleri"] = []
        form_data["harcama_bildirimleri"] = []
        form_data["last_step"] = 0
        form_data["assigned_to_user_id"] = None
        form_data["assigned_by_user_id"] = None
        form_data["assigned_at"] = None

        if linked_request:
            description_parts = ["Müşteri Talebi:", linked_request.get("request_description", "")]
            requirements = linked_request.get("requirements")
            if requirements:
                description_parts.extend(["", "Gereklilikler:", requirements])
            form_data["gorev_tanimi"] = "\n".join(part for part in description_parts if part is not None)
            if linked_request.get("customer_address"):
                form_data["gorev_yeri"] = linked_request["customer_address"]
            if linked_request.get("customer_name"):
                form_data["gorev_firma"] = linked_request["customer_name"]

        store_form_in_session(form_no, form_data)
        if linked_request:
            try:
                task_request_service.mark_converted(
                    linked_request["id"], form_no=form_no, base_path=str(BASE_PATH)
                )
            except task_request_service.TaskRequestError as exc:
                flash(str(exc), "error")
                return redirect(url_for("task_requests_list"))
            flash(f"Talep göreve dönüştürüldü: {form_no}", "success")
        else:
            flash(f"Yeni form oluşturuldu. Form No: {form_no}", "success")
        redirect_url = url_for("form_wizard", form_no=form_no, step=0)
        print(f"🔍 Oluşturulan form_no: '{form_no}'")
        print(f"🔍 Redirect URL: {redirect_url}")
        return redirect(redirect_url)

    @app.route("/form/<form_no>", methods=["GET", "POST"])
    def form_wizard(form_no: str):
        response = require_login()
        if response is not None:
            return response

        step = clamp_step(request.values.get("step", 0))
        form_data = ensure_form_data(form_no)
        if form_data is None:
            flash(f"Form {form_no} yüklenemedi.", "error")
            return redirect(url_for("index"))

        current = get_current_user()
        role = current.get("role") if current else None
        is_employee = role == "calisan"
        if is_employee and form_data.get("assigned_to_user_id") != current.get("id"):
            flash("Bu göreve erişiminiz yok.", "error")
            return redirect(url_for("index"))

        if is_form_locked(form_no):
            flash("Tamamlanmış formlar düzenlenemez. Özet sayfasına yönlendirildiniz.", "warning")
            return redirect(url_for("form_summary", form_no=form_no))

        current_step = FORM_STEPS[step]
        read_only_step = is_employee and current_step["id"] != "gorev_bilgileri"

        if request.method == "POST":
            if read_only_step:
                flash("Bu adımda değişiklik yapamazsınız.", "warning")
                return redirect(url_for("form_wizard", form_no=form_no, step=step))

            action = request.form.get("action", "next")
            remove_target = ""
            remove_expense_index = ""
            if current_step["id"] == "gorev_bilgileri":
                remove_target = request.form.get("remove_attachment", "").strip()
                remove_expense_index = request.form.get("remove_expense", "").strip()

            result = update_form_data_from_request(
                form_no,
                form_data,
                current_step["id"],
                request.form,
                request.files,
                action=action,
            )
            attachments = normalize_attachments(form_data)
            expenses = form_data.get("harcama_bildirimleri", [])

            if remove_target:
                remaining = [item for item in attachments if item["filename"] != remove_target]
                if len(remaining) != len(attachments):
                    form_data["gorev_ekleri"] = remaining
                    file_deleted = delete_attachment_file(form_no, remove_target)
                    message = "Ek kaldırıldı."
                    category = "info"
                    if not file_deleted:
                        message = "Ek listeden kaldırıldı ancak dosya bulunamadı."
                        category = "warning"
                    flash(message, category)
                else:
                    flash("Silinecek ek bulunamadı.", "warning")
                update_last_step(form_data, step)
                store_form_in_session(form_no, form_data)
                return redirect(url_for("form_wizard", form_no=form_no, step=step))

            if action == "remove_expense":
                index = None
                if remove_expense_index.isdigit():
                    index = int(remove_expense_index)
                if index is not None and 0 <= index < len(expenses):
                    removed = expenses.pop(index)
                    form_data["harcama_bildirimleri"] = expenses
                    for attachment in removed.get("attachments", []):
                        delete_attachment_file(form_no, attachment.get("filename", ""))
                    flash("Harcama bildirimi kaldırıldı.", "info")
                else:
                    flash("Silinecek harcama bildirimi bulunamadı.", "warning")
                update_last_step(form_data, step)
                store_form_in_session(form_no, form_data)
                return redirect(url_for("form_wizard", form_no=form_no, step=step))

            if action == "add_expense":
                if result.get("expense_error"):
                    flash(result["expense_error"], "warning")
                elif result.get("expense_added"):
                    flash("Harcama bildirimi eklendi.", "success")
                else:
                    flash("Harcama bildirimi eklenemedi.", "warning")
                update_last_step(form_data, step)
                store_form_in_session(form_no, form_data)
                return redirect(url_for("form_wizard", form_no=form_no, step=step))

            target_step = step
            if action == "previous":
                target_step = max(0, step - 1)
            elif action == "save":
                target_step = step
            else:
                target_step = min(total_steps - 1, step + 1)

            if is_employee:
                target_step = max(total_steps - 1, target_step)

            update_last_step(form_data, target_step)
            store_form_in_session(form_no, form_data)

            if action == "previous":
                previous_step = max(0, step - 1)
                return redirect(url_for("form_wizard", form_no=form_no, step=previous_step))
            if action == "save":
                try:
                    _, status = form_service.save_form(form_no, form_data, base_path=str(BASE_PATH))
                except FormServiceError as exc:
                    flash(str(exc), "error")
                else:
                    form_data["durum"] = status.code
                    store_form_in_session(form_no, form_data)
                    flash(f"Form {status.code} olarak veritabanına kaydedildi.", "success")
                return redirect(url_for("form_wizard", form_no=form_no, step=step))

            next_step = step + 1
            if next_step >= total_steps:
                return redirect(url_for("form_summary", form_no=form_no))
            return redirect(url_for("form_wizard", form_no=form_no, step=next_step))

        dynamic_data = get_dynamic_data()
        assigned_user = None
        assigned_by_user = None
        if form_data.get("assigned_to_user_id"):
            assigned_user = user_service.get_user(
                int(form_data["assigned_to_user_id"]), base_path=str(BASE_PATH)
            )
        if form_data.get("assigned_by_user_id"):
            assigned_by_user = user_service.get_user(
                int(form_data["assigned_by_user_id"]), base_path=str(BASE_PATH)
            )

        personel_users = user_service.list_users_by_role(
            "calisan", base_path=str(BASE_PATH)
        )
        hazirlayan_users = user_service.list_users_by_roles(
            ("admin", "atayan"), base_path=str(BASE_PATH)
        )

        responsible_name = ""
        if assigned_user:
            responsible_name = assigned_user.full_name
        elif form_data.get("personel_1"):
            responsible_name = form_data["personel_1"]
        return render_template(
            current_step["template"],
            form_no=form_no,
            form_data=form_data,
            step_index=step,
            taseron_options=dynamic_data["taseron_options"],
            arac_plaka_options=dynamic_data["arac_plaka_options"],
            hazirlayan_users=hazirlayan_users,
            personel_users=personel_users,
            read_only_step=read_only_step,
            is_employee=is_employee,
            assigned_user=assigned_user,
            assigned_by_user=assigned_by_user,
            responsible_name=responsible_name,
        )

    @app.route("/form/<form_no>/summary", methods=["GET", "POST"])
    def form_summary(form_no: str):
        response = require_login()
        if response is not None:
            return response

        form_data = ensure_form_data(form_no)
        if form_data is None:
            flash(f"Form {form_no} yüklenemedi.", "error")
            return redirect(url_for("index"))

        current = get_current_user()
        is_employee = current and current.get("role") == "calisan"
        is_responsible = bool(
            is_employee and form_data.get("assigned_to_user_id") == current.get("id")
        )
        is_team_member = bool(is_employee and user_is_form_personnel(current, form_data))

        if is_employee and not (is_responsible or is_team_member):
            flash("Bu göreve erişiminiz yok.", "error")
            return redirect(url_for("index"))

        status = form_service.determine_form_status(form_data)
        form_data["durum"] = status.code
        store_form_in_session(form_no, form_data)

        locked = is_form_locked(form_no)
        can_edit = not locked and (not is_employee or is_responsible)

        if request.method == "POST":
            if locked:
                flash("Tamamlanmış formlar üzerinde değişiklik yapılamaz.", "warning")
                return redirect(url_for("form_summary", form_no=form_no))
            action = request.form.get("action")
            if action == "save":
                try:
                    update_last_step(form_data, total_steps - 1)
                    store_form_in_session(form_no, form_data)
                    _, status = form_service.save_form(form_no, form_data, base_path=str(BASE_PATH))
                except FormServiceError as exc:
                    flash(str(exc), "error")
                else:
                    form_data["durum"] = status.code
                    store_form_in_session(form_no, form_data)
                    flash(f"Form {status.code} olarak veritabanına kaydedildi.", "success")
                return redirect(url_for("form_summary", form_no=form_no))
            if action == "previous":
                return redirect(url_for("form_wizard", form_no=form_no, step=total_steps - 1))

        missing_fields = [
            FIELD_LABELS.get(field, field.replace("_", " ").title())
            for field in status.missing_fields
        ]

        assigned_user = None
        assigned_by_user = None
        if form_data.get("assigned_to_user_id"):
            assigned_user = user_service.get_user(
                int(form_data["assigned_to_user_id"]), base_path=str(BASE_PATH)
            )
        if form_data.get("assigned_by_user_id"):
            assigned_by_user = user_service.get_user(
                int(form_data["assigned_by_user_id"]), base_path=str(BASE_PATH)
            )

        return render_template(
            "summary.html",
            form_no=form_no,
            form_data=form_data,
            status=status,
            missing_fields=missing_fields,
            locked=locked,
            can_edit=can_edit,
            assigned_user=assigned_user,
            assigned_by_user=assigned_by_user,
        )

    @app.get("/form/<form_no>/attachments/<path:filename>")
    def download_attachment(form_no: str, filename: str):
        form_data = ensure_form_data(form_no)
        if form_data is None:
            flash(f"Form {form_no} yüklenemedi.", "error")
            return redirect(url_for("index"))

        attachments = list(form_data.get("gorev_ekleri") or [])
        for entry in form_data.get("harcama_bildirimleri", []):
            if isinstance(entry, dict):
                for attachment in entry.get("attachments", []) or []:
                    if isinstance(attachment, dict):
                        attachments.append(attachment)
        original_name = next(
            (item.get("original_name") for item in attachments if item.get("filename") == filename),
            filename,
        )

        base_dir = (UPLOAD_DIR / form_no).resolve()
        try:
            file_path = (base_dir / filename).resolve()
            file_path.relative_to(base_dir)
        except (ValueError, RuntimeError):
            flash("Dosya bulunamadı.", "error")
            return redirect(url_for("form_summary", form_no=form_no))

        if not file_path.exists() or not file_path.is_file():
            flash("Dosya bulunamadı.", "error")
            return redirect(url_for("form_summary", form_no=form_no))

        return send_from_directory(
            file_path.parent,
            file_path.name,
            as_attachment=True,
            download_name=original_name,
        )

    @app.route("/admin")
    def admin_panel():
        response = ensure_admin_access()
        if response is not None:
            return response

        dynamic_data = get_dynamic_data()
        form_defaults = get_form_defaults()
        admin_users = user_service.list_users_by_role("admin", base_path=str(BASE_PATH))
        assign_users = user_service.list_users_by_role("atayan", base_path=str(BASE_PATH))
        employee_users = user_service.list_users_by_role("calisan", base_path=str(BASE_PATH))
        return render_template(
            "admin.html",
            data=dynamic_data,
            form_defaults=form_defaults,
            admin_users=admin_users,
            assign_users=assign_users,
            employee_users=employee_users,
        )

    @app.post("/admin/update")
    def admin_update():
        response = ensure_admin_access()
        if response is not None:
            return response

        current_data = get_dynamic_data()
        updated_data: Dict[str, List[str]] = {}
        for key in current_data:
            raw_value = request.form.get(key, "")
            lines = [line.strip() for line in raw_value.splitlines()]
            updated_data[key] = [line for line in lines if line]

        set_dynamic_data(updated_data)

        set_form_defaults(
            {
                "dok_no": request.form.get("default_dok_no", ""),
                "rev_no": request.form.get("default_rev_no", ""),
            }
        )
        flash("Veri listeleri güncellendi.", "success")
        return redirect(url_for("admin_panel"))

    @app.post("/admin/users/<role>/create")
    def admin_create_user(role: str):
        response = ensure_admin_access()
        if response is not None:
            return response

        normalized_role = (role or "").strip().lower()
        if normalized_role not in {"admin", "atayan", "calisan"}:
            flash("Geçersiz rol seçimi.", "error")
            return redirect(url_for("admin_panel"))

        full_name = request.form.get("full_name", "")
        email = request.form.get("email")
        phone = request.form.get("phone")
        password = request.form.get("password")

        try:
            created = user_service.create_user(
                full_name=full_name,
                email=email,
                phone=phone,
                password=password,
                role=normalized_role,
                base_path=str(BASE_PATH),
            )
        except UserServiceError as exc:
            flash(str(exc), "error")
        else:
            flash(f"{created.full_name} başarıyla eklendi.", "success")
        return redirect(url_for("admin_panel"))

    @app.post("/admin/users/<int:user_id>/delete")
    def admin_delete_user(user_id: int):
        response = ensure_admin_access()
        if response is not None:
            return response

        current = get_current_user()
        if current and current.get("id") == user_id:
            flash("Kendi hesabınızı silemezsiniz.", "warning")
            return redirect(url_for("admin_panel"))

        user_obj = user_service.get_user(user_id, base_path=str(BASE_PATH))
        if user_obj is None:
            flash("Kullanıcı bulunamadı.", "error")
            return redirect(url_for("admin_panel"))

        user_service.delete_user(user_id, base_path=str(BASE_PATH))
        flash(f"{user_obj.full_name} silindi.", "info")
        return redirect(url_for("admin_panel"))

    @app.post("/admin/upload")
    def admin_upload():
        response = ensure_admin_access()
        if response is not None:
            return response

        data_file = request.files.get("data_file")
        if data_file is None or not data_file.filename:
            flash("Lütfen bir JSON dosyası seçin.", "warning")
            return redirect(url_for("admin_panel"))

        try:
            payload = json.loads(data_file.read().decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            flash("JSON dosyası okunamadı.", "error")
            return redirect(url_for("admin_panel"))

        if not isinstance(payload, dict):
            flash("Geçersiz JSON formatı.", "error")
            return redirect(url_for("admin_panel"))

        set_dynamic_data({key: payload.get(key, []) for key in DEFAULT_DYNAMIC_DATA})
        migrate_legacy_user_lists(
            base_path=str(BASE_PATH),
            assigners=payload.get("hazirlayan_options"),
            employees=payload.get("personel_options"),
        )

        defaults_payload = payload.get("form_defaults", {})
        if isinstance(defaults_payload, dict):
            set_form_defaults(defaults_payload)
        else:
            set_form_defaults({})
        flash("Yeni JSON dosyası yüklendi.", "success")
        return redirect(url_for("admin_panel"))

    @app.get("/logout")
    def logout():
        session.clear()
        flash("Çıkış yapıldı. Tekrar görüşmek üzere!", "info")
        return redirect(url_for("index"))

    @app.get("/admin/logout")
    def admin_logout():
        return redirect(url_for("logout"))

    @app.get("/form/<form_no>/export/excel")
    def export_form_excel(form_no: str):
        form_data = ensure_form_data(form_no)
        if form_data is None:
            flash(f"Form {form_no} bulunamadı.", "error")
            return redirect(url_for("index"))

        stream = form_service.export_form_to_excel(form_no, form_data)
        filename = f"gorev_formu_{form_no}.xlsx"
        return send_file(
            stream,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.get("/form/<form_no>/export/pdf")
    def export_form_pdf(form_no: str):
        form_data = ensure_form_data(form_no)
        if form_data is None:
            flash(f"Form {form_no} bulunamadı.", "error")
            return redirect(url_for("index"))

        stream = form_service.export_form_to_pdf(form_no, form_data)
        filename = f"gorev_formu_{form_no}.pdf"
        return send_file(
            stream,
            as_attachment=True,
            download_name=filename,
            mimetype="application/pdf",
        )

    def clamp_step(step_value) -> int:
        try:
            value = int(step_value)
        except (ValueError, TypeError):
            return 0
        return max(0, min(value, total_steps - 1))

    def update_last_step(form_data: Dict[str, Any], value: int) -> None:
        form_data["last_step"] = clamp_step(value)

    def ensure_form_data(form_no: str):
        forms = session.get("forms", {})
        form_data = forms.get(form_no)
        if form_data:
            normalize_attachments(form_data)
            update_last_step(form_data, form_data.get("last_step", 0))
            return form_data
        try:
            loaded = form_service.load_form_data(form_no, base_path=str(BASE_PATH))
        except FormServiceError as exc:
            flash(str(exc), "error")
            return None
        else:
            normalize_attachments(loaded)
            store_form_in_session(form_no, loaded)
            return loaded

    def store_form_in_session(form_no: str, form_data: Dict[str, Any]) -> None:
        normalize_attachments(form_data)
        update_last_step(form_data, form_data.get("last_step", 0))
        forms = session.get("forms", {})
        forms[form_no] = form_data
        session["forms"] = forms

    def normalize_date(value: str) -> str:
        value = (value or "").strip()
        if not value:
            return ""
        for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
            try:
                parsed = datetime.strptime(value, fmt)
            except ValueError:
                continue
            else:
                return parsed.strftime("%d.%m.%Y")
        return value

    def normalize_time(value: str) -> str:
        value = (value or "").strip()
        if not value:
            return ""
        try:
            return datetime.strptime(value, "%H:%M").strftime("%H:%M")
        except ValueError:
            return value

    def update_form_data_from_request(
        form_no: str,
        form_data: Dict[str, Any],
        step_id: str,
        data,
        files,
        *,
        action: str = "",
    ) -> Dict[str, Any]:
        result: Dict[str, Any] = {}
        if step_id == "form_bilgileri":
            form_data["dok_no"] = data.get("dok_no", "").strip()
            form_data["rev_no"] = data.get("rev_no", "").strip()
            tarih = data.get("tarih", "").strip()
            if tarih:
                form_data["tarih"] = tarih
        elif step_id == "hazirlayan":
            selected = data.get("hazirlayan", "").strip()
            valid_assigners = {
                user.full_name for user in user_service.list_users_by_roles(("admin", "atayan"), base_path=str(BASE_PATH))
            }
            form_data["hazirlayan"] = selected if selected in valid_assigners else ""
        elif step_id == "gorevli_personel":
            valid_employees = {
                user.full_name for user in user_service.list_users_by_role("calisan", base_path=str(BASE_PATH))
            }
            for index in range(1, 6):
                value = data.get(f"personel_{index}", "").strip()
                form_data[f"personel_{index}"] = value if value in valid_employees else ""
            form_data["gorev_tarih"] = normalize_date(data.get("gorev_tarih", ""))
            primary_employee = form_data.get("personel_1", "").strip()
            if primary_employee:
                employee_obj = user_service.get_user_by_name(
                    primary_employee, base_path=str(BASE_PATH)
                )
                if employee_obj and employee_obj.role == "calisan":
                    current_user_obj = get_current_user() or {}
                    form_data["assigned_to_user_id"] = employee_obj.id
                    form_data["assigned_by_user_id"] = current_user_obj.get("id")
                    form_data["assigned_at"] = datetime.now().isoformat()
                else:
                    form_data["assigned_to_user_id"] = None
                    form_data["assigned_by_user_id"] = None
                    form_data["assigned_at"] = None
            else:
                form_data["assigned_to_user_id"] = None
                form_data["assigned_by_user_id"] = None
                form_data["assigned_at"] = None
        elif step_id == "finans_arac":
            form_data["avans"] = data.get("avans", "").strip()
            form_data["taseron"] = data.get("taseron", "").strip()
            form_data["arac_plaka"] = data.get("arac_plaka", "").strip()
        elif step_id == "gorev_detay":
            form_data["gorev_tanimi"] = data.get("gorev_tanimi", "").strip()
            form_data["gorev_yeri"] = data.get("gorev_yeri", "").strip()
            form_data["gorev_il"] = data.get("gorev_il", "").strip()
            form_data["gorev_ilce"] = data.get("gorev_ilce", "").strip()
            form_data["gorev_firma"] = data.get("gorev_firma", "").strip()
        elif step_id == "gorev_bilgileri":
            date_fields = [
                "yola_cikis_tarih",
                "donus_tarih",
                "calisma_baslangic_tarih",
                "calisma_bitis_tarih",
            ]
            time_fields = [
                "yola_cikis_saat",
                "donus_saat",
                "calisma_baslangic_saat",
                "calisma_bitis_saat",
            ]

            for field in date_fields:
                form_data[field] = normalize_date(data.get(field, ""))
            for field in time_fields:
                form_data[field] = normalize_time(data.get(field, ""))
            form_data["mola_suresi"] = data.get("mola_suresi", "").strip()
            form_data["yapilan_isler"] = data.get("yapilan_isler", "").strip()
            existing = form_data.get("gorev_ekleri")
            if not isinstance(existing, list):
                existing = []
            uploads = save_uploaded_files(form_no, files, "gorev_ekleri")
            if uploads:
                existing.extend(uploads)
            form_data["gorev_ekleri"] = existing
            expenses = form_data.get("harcama_bildirimleri")
            if not isinstance(expenses, list):
                expenses = []
            if action == "add_expense":
                description = data.get("harcama_aciklamasi", "").strip()
                receipt_uploads = save_uploaded_files(form_no, files, "harcama_dosyalari")
                if description or receipt_uploads:
                    expenses.append(
                        {
                            "description": description,
                            "attachments": receipt_uploads,
                        }
                    )
                    result["expense_added"] = True
                else:
                    result["expense_error"] = "Harcama eklemek için açıklama veya görsel yükleyin."
            form_data["harcama_bildirimleri"] = expenses
        return result
    app.jinja_env.globals.update(
        FIELD_LABELS=FIELD_LABELS,
        datetime=datetime,
    )


app = create_app()


__all__ = ["app", "create_app"]
