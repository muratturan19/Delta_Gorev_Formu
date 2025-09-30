from __future__ import annotations

import glob
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, List

from flask import (Flask, flash, redirect, render_template, request, session,
                   url_for)

from core import form_service
from core.form_service import FormServiceError

BASE_PATH = Path(__file__).resolve().parents[1]
DEFAULT_FORM_VALUES: Dict[str, str] = {
    "dok_no": "F-001",
    "rev_no": "",
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

PERSONEL_OPTIONS: List[str] = [
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
]

TASERON_OPTIONS: List[str] = [
    "Yok",
    "ABC İnşaat",
    "XYZ Teknik",
    "Marmara Mühendislik",
    "Anadolu Yapı",
]

HAZIRLAYAN_OPTIONS: List[str] = [
    "Ali Yılmaz",
    "Ayşe Demir",
    "Mehmet Korkmaz",
    "Zeynep Ak",
    "Elif Kaya",
]

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


def create_app() -> Flask:
    app = Flask(__name__)
    app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-key")

    register_routes(app)
    return app


def register_routes(app: Flask) -> None:
    total_steps = len(FORM_STEPS)

    @app.context_processor
    def inject_globals():  # pragma: no cover - template helper
        return {
            "FORM_STEPS": FORM_STEPS,
            "total_steps": total_steps,
        }

    @app.route("/")
    def index():
        excel_files = glob.glob(str(BASE_PATH / "gorev_formu_*.xlsx"))
        form_numbers = sorted(
            [Path(file).stem.replace("gorev_formu_", "") for file in excel_files],
            reverse=True,
        )
        return render_template("home.html", form_numbers=form_numbers)

    @app.post("/form/load")
    def load_form_redirect():
        form_no = request.form.get("form_no_input", "").strip() or request.form.get("form_no_select", "").strip()
        if not form_no:
            flash("Form numarası seçin veya girin.", "warning")
            return redirect(url_for("index"))
        return redirect(url_for("form_wizard", form_no=form_no, step=0))

    @app.route("/form/new")
    def new_form():
        try:
            form_no = form_service.get_next_form_no(base_path=str(BASE_PATH))
        except FormServiceError as exc:
            flash(str(exc), "error")
            return redirect(url_for("index"))

        form_data = {
            **DEFAULT_FORM_VALUES,
            "form_no": form_no,
            "tarih": datetime.now().strftime("%d.%m.%Y"),
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
                    filename, status = form_service.save_form(form_no, form_data, base_path=str(BASE_PATH))
                except FormServiceError as exc:
                    flash(str(exc), "error")
                else:
                    form_data["durum"] = status.code
                    store_form_in_session(form_no, form_data)
                    flash(f"Form {status.code} olarak kaydedildi. Dosya: {filename}", "success")
                return redirect(url_for("form_wizard", form_no=form_no, step=step))

            next_step = step + 1
            if next_step >= total_steps:
                return redirect(url_for("form_summary", form_no=form_no))
            return redirect(url_for("form_wizard", form_no=form_no, step=next_step))

        return render_template(
            current_step["template"],
            form_no=form_no,
            form_data=form_data,
            step_index=step,
            personel_options=PERSONEL_OPTIONS,
            taseron_options=TASERON_OPTIONS,
            hazirlayan_options=HAZIRLAYAN_OPTIONS,
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

        if request.method == "POST":
            action = request.form.get("action")
            if action == "save":
                try:
                    filename, status = form_service.save_form(form_no, form_data, base_path=str(BASE_PATH))
                except FormServiceError as exc:
                    flash(str(exc), "error")
                else:
                    form_data["durum"] = status.code
                    store_form_in_session(form_no, form_data)
                    flash(f"Form {status.code} olarak kaydedildi. Dosya: {filename}", "success")
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
            for field in [
                "yola_cikis_tarih",
                "yola_cikis_saat",
                "donus_tarih",
                "donus_saat",
                "calisma_baslangic_tarih",
                "calisma_baslangic_saat",
                "calisma_bitis_tarih",
                "calisma_bitis_saat",
                "mola_suresi",
            ]:
                form_data[field] = data.get(field, "").strip()
        elif step_id == "arac_bilgisi":
            form_data["arac_plaka"] = data.get("arac_plaka", "").strip()
        elif step_id == "hazirlayan":
            form_data["hazirlayan"] = data.get("hazirlayan", "").strip()

    app.jinja_env.globals.update(
        PERSONEL_OPTIONS=PERSONEL_OPTIONS,
        TASERON_OPTIONS=TASERON_OPTIONS,
        HAZIRLAYAN_OPTIONS=HAZIRLAYAN_OPTIONS,
        FIELD_LABELS=FIELD_LABELS,
        datetime=datetime,
    )


app = create_app()


__all__ = ["app", "create_app"]
