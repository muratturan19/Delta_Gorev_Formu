# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

from flask import (Flask, flash, redirect, render_template, request, send_file,
                   session, url_for)

from core import form_service
from core.form_service import FormServiceError

BASE_PATH = Path(__file__).resolve().parents[1]
DATA_FILE = BASE_PATH / "data.json"
DEFAULT_FORM_VALUES: Dict[str, str] = {
    "dok_no": "F-001",
    "rev_no": "00 / 06.05.24",
    "avans": "",
    "taseron": "",
    "gorev_tanimi": "",
    "gorev_yeri": "",
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
    "durum": "YARIM",
}
DEFAULT_FORM_VALUES.update({f"personel_{index}": "" for index in range(1, 6)})

FIELD_LABELS: Dict[str, str] = {
    "dok_no": "DOK.NO",
    "rev_no": "REV.NO/TRH",
    "avans": "Avans Tutarı",
    "taseron": "Taşeron Şirket",
    "gorev_tanimi": "Görevin Tanımı",
    "gorev_yeri": "Görev Yeri",
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
}

DEFAULT_DYNAMIC_DATA: Dict[str, List[str]] = {
    "personel_options": [
        "Ahmet Yılmaz",
        "Mehmet Demir",
        "Ali Kaya",
        "Veli Çelik",
        "Hasan Şahin",
        "Hüseyin Aydın",
        "İbrahim Özdemir",
        "Mustafa Arslan",
        "Emre Doğan",
        "Burak Yıldız",
    ],
    "taseron_options": [
        "Yok",
        "ABC İnşaat",
        "XYZ Teknik",
        "Marmara Mühendislik",
        "Anadolu Yapı",
    ],
    "hazirlayan_options": [
        "Ali Yılmaz",
        "Ayşe Demir",
        "Mehmet Korkmaz",
        "Zeynep Ak",
        "Elif Kaya",
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

_storage: Dict[str, Any] | None = None
_dynamic_data: Dict[str, List[str]] | None = None
_form_defaults: Dict[str, str] | None = None
ADMIN_PASSWORD = os.environ.get("ADMIN_PANEL_PASSWORD") or os.environ.get("ADMIN_PASSWORD") or "delta-admin"

FORM_STEPS: List[Dict[str, str]] = [
    {"id": "form_bilgileri", "title": "Form Bilgileri", "template": "steps/form_bilgileri.html"},
    {"id": "gorevli_personel", "title": "Görevli Personel", "template": "steps/gorevli_personel.html"},
    {"id": "avans_taseron", "title": "Avans ve Taşeron", "template": "steps/avans_taseron.html"},
    {"id": "gorev_tanimi", "title": "Görev Tanımı", "template": "steps/gorev_tanimi.html"},
    {"id": "gorev_yeri", "title": "Görev Yeri", "template": "steps/gorev_yeri.html"},
    {"id": "saat_bilgileri", "title": "Saat Bilgileri", "template": "steps/saat_bilgileri.html"},
    {"id": "arac_bilgisi", "title": "Araç Bilgisi", "template": "steps/arac_bilgisi.html"},
    {"id": "hazirlayan", "title": "Hazırlayan", "template": "steps/hazirlayan.html"},
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
    app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-key")

    register_routes(app)
    return app


def register_routes(app: Flask) -> None:
    total_steps = len(FORM_STEPS)

    def is_admin() -> bool:
        return bool(session.get("is_admin"))

    def ensure_admin_access():
        if not is_admin():
            flash("Admin paneline erişmek için giriş yapın.", "error")
            return redirect(url_for("admin_panel"))
        return None

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
    def inject_globals():  # pragma: no cover - template helper
        dynamic_data = get_dynamic_data()
        return {
            "FORM_STEPS": FORM_STEPS,
            "total_steps": total_steps,
            "personel_options": dynamic_data["personel_options"],
            "taseron_options": dynamic_data["taseron_options"],
            "hazirlayan_options": dynamic_data["hazirlayan_options"],
            "arac_plaka_options": dynamic_data["arac_plaka_options"],
            "is_admin": is_admin(),
        }

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

    @app.route("/")
    def index():
        form_numbers = form_service.list_form_numbers(base_path=str(BASE_PATH))

        filters = {
            "personel": request.args.get("personel", "").strip(),
            "gorev_yeri": request.args.get("gorev_yeri", "").strip(),
            "start_date": request.args.get("start_date", "").strip(),
            "end_date": request.args.get("end_date", "").strip(),
        }

        performed_search = any(filters.values())
        search_results = []
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
        )

    @app.post("/form/load")
    def load_form_redirect():
        form_no = request.form.get("form_no_input", "").strip() or request.form.get("form_no_select", "").strip()
        if not form_no:
            flash("Form numarası seçin veya girin.", "warning")
            return redirect(url_for("index"))
        form_data = ensure_form_data(form_no)
        if form_data is None:
            return redirect(url_for("index"))

        status = form_service.determine_form_status(form_data)
        form_data["durum"] = status.code
        store_form_in_session(form_no, form_data)

        if status.is_complete:
            lock_form(form_no)
            flash("Tamamlanmış form özet görünümünde açıldı.", "info")
            return redirect(url_for("form_summary", form_no=form_no))

        unlock_form(form_no)
        return redirect(url_for("form_wizard", form_no=form_no, step=0))

    @app.route("/form/new")
    def new_form():
        try:
            form_no = form_service.get_next_form_no(base_path=str(BASE_PATH))
        except FormServiceError as exc:
            flash(str(exc), "error")
            return redirect(url_for("index"))

        form_defaults = get_form_defaults()
        form_data = {
            **DEFAULT_FORM_VALUES,
            "form_no": form_no,
            "tarih": datetime.now().strftime("%d.%m.%Y"),
            "dok_no": form_defaults["dok_no"],
            "rev_no": form_defaults["rev_no"],
        }
        store_form_in_session(form_no, form_data)
        flash(f"Yeni form oluşturuldu. Form No: {form_no}", "success")
        return redirect(url_for("form_wizard", form_no=form_no, step=0))

    @app.route("/form/<form_no>", methods=["GET", "POST"])
    def form_wizard(form_no: str):
        step = clamp_step(request.values.get("step", 0))
        form_data = ensure_form_data(form_no)
        if form_data is None:
            flash(f"Form {form_no} yüklenemedi.", "error")
            return redirect(url_for("index"))

        if is_form_locked(form_no):
            flash("Tamamlanmış formlar düzenlenemez. Özet sayfasına yönlendirildiniz.", "warning")
            return redirect(url_for("form_summary", form_no=form_no))

        current_step = FORM_STEPS[step]

        if request.method == "POST":
            update_form_data_from_request(form_data, current_step["id"], request.form)
            store_form_in_session(form_no, form_data)

            action = request.form.get("action")
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
        return render_template(
            current_step["template"],
            form_no=form_no,
            form_data=form_data,
            step_index=step,
            personel_options=dynamic_data["personel_options"],
            taseron_options=dynamic_data["taseron_options"],
            hazirlayan_options=dynamic_data["hazirlayan_options"],
            arac_plaka_options=dynamic_data["arac_plaka_options"],
        )

    @app.route("/form/<form_no>/summary", methods=["GET", "POST"])
    def form_summary(form_no: str):
        form_data = ensure_form_data(form_no)
        if form_data is None:
            flash(f"Form {form_no} yüklenemedi.", "error")
            return redirect(url_for("index"))

        status = form_service.determine_form_status(form_data)
        form_data["durum"] = status.code
        store_form_in_session(form_no, form_data)

        locked = is_form_locked(form_no)

        if request.method == "POST":
            if locked:
                flash("Tamamlanmış formlar üzerinde değişiklik yapılamaz.", "warning")
                return redirect(url_for("form_summary", form_no=form_no))
            action = request.form.get("action")
            if action == "save":
                try:
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

        return render_template(
            "summary.html",
            form_no=form_no,
            form_data=form_data,
            status=status,
            missing_fields=missing_fields,
            locked=locked,
        )

    @app.route("/admin", methods=["GET", "POST"])
    def admin_panel():
        if not is_admin():
            if request.method == "POST":
                password = request.form.get("password", "")
                if password == ADMIN_PASSWORD:
                    session["is_admin"] = True
                    flash("Admin paneline giriş yapıldı.", "success")
                    return redirect(url_for("admin_panel"))
                flash("Hatalı şifre.", "error")
            return render_template("admin_login.html")

        dynamic_data = get_dynamic_data()
        form_defaults = get_form_defaults()
        return render_template("admin.html", data=dynamic_data, form_defaults=form_defaults)

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

        defaults_payload = payload.get("form_defaults", {})
        if isinstance(defaults_payload, dict):
            set_form_defaults(defaults_payload)
        else:
            set_form_defaults({})
        flash("Yeni JSON dosyası yüklendi.", "success")
        return redirect(url_for("admin_panel"))

    @app.get("/admin/logout")
    def admin_logout():
        session.pop("is_admin", None)
        flash("Admin oturumu kapatıldı.", "info")
        return redirect(url_for("index"))

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

    def ensure_form_data(form_no: str):
        forms = session.get("forms", {})
        form_data = forms.get(form_no)
        if form_data:
            return form_data
        try:
            loaded = form_service.load_form_data(form_no, base_path=str(BASE_PATH))
        except FormServiceError as exc:
            flash(str(exc), "error")
            return None
        else:
            store_form_in_session(form_no, loaded)
            return loaded

    def store_form_in_session(form_no: str, form_data: Dict[str, str]) -> None:
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

    def update_form_data_from_request(form_data: Dict[str, str], step_id: str, data) -> None:
        if step_id == "form_bilgileri":
            form_data["dok_no"] = data.get("dok_no", "").strip()
            form_data["rev_no"] = data.get("rev_no", "").strip()
            tarih = data.get("tarih", "").strip()
            if tarih:
                form_data["tarih"] = tarih
        elif step_id == "gorevli_personel":
            for index in range(1, 6):
                form_data[f"personel_{index}"] = data.get(f"personel_{index}", "").strip()
        elif step_id == "avans_taseron":
            form_data["avans"] = data.get("avans", "").strip()
            form_data["taseron"] = data.get("taseron", "").strip()
        elif step_id == "gorev_tanimi":
            form_data["gorev_tanimi"] = data.get("gorev_tanimi", "").strip()
        elif step_id == "gorev_yeri":
            form_data["gorev_yeri"] = data.get("gorev_yeri", "").strip()
        elif step_id == "saat_bilgileri":
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
        elif step_id == "arac_bilgisi":
            form_data["arac_plaka"] = data.get("arac_plaka", "").strip()
        elif step_id == "hazirlayan":
            form_data["hazirlayan"] = data.get("hazirlayan", "").strip()

    app.jinja_env.globals.update(
        FIELD_LABELS=FIELD_LABELS,
        datetime=datetime,
    )


app = create_app()


__all__ = ["app", "create_app"]
