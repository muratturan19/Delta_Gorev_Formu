"""Servis katmanı: Görev formu veri işlemleri."""
from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side


CONFIG_FILE_NAME = "form_config.json"
FORM_FILENAME_TEMPLATE = "gorev_formu_{form_no}.xlsx"


class FormServiceError(Exception):
    """Servis katmanına özgü hata sınıfı."""


@dataclass
class FormStatus:
    """Form durumunu ve eksik alanları temsil eder."""

    code: str
    missing_fields: List[str]

    @property
    def is_complete(self) -> bool:
        return self.code.upper() == "TAMAMLANDI"


def get_next_form_no(base_path: str = ".") -> str:
    """Konfigürasyonda saklanan bir sonraki form numarasını döndür."""

    config_path = os.path.join(base_path, CONFIG_FILE_NAME)
    last_no = 0

    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as file:
            try:
                config = json.load(file)
            except json.JSONDecodeError as exc:  # Bozuk dosya varsa sıfırla
                raise FormServiceError("Konfigürasyon dosyası okunamadı.") from exc
            last_no = int(config.get("last_form_no", 0))

    next_no = last_no + 1
    with open(config_path, "w", encoding="utf-8") as file:
        json.dump({"last_form_no": next_no}, file)

    return str(next_no).zfill(5)


def get_excel_filename(form_no: str, base_path: str = ".") -> str:
    """Verilen form numarası için dosya yolunu döndür."""

    filename = FORM_FILENAME_TEMPLATE.format(form_no=form_no)
    return os.path.join(base_path, filename)


def determine_form_status(form_data: Dict[str, Any]) -> FormStatus:
    """Formun tamamlanma durumunu belirle."""

    required_fields = [
        "yola_cikis_tarih",
        "yola_cikis_saat",
        "calisma_baslangic_tarih",
        "calisma_baslangic_saat",
        "calisma_bitis_tarih",
        "calisma_bitis_saat",
        "donus_tarih",
        "donus_saat",
    ]

    missing_fields: List[str] = []
    for key in required_fields:
        value = (form_data.get(key) or "").strip()
        if not value:
            missing_fields.append(key)

    status_code = "TAMAMLANDI" if not missing_fields else "YARIM"
    return FormStatus(code=status_code, missing_fields=missing_fields)


def load_form_data(form_no: str, base_path: str = ".") -> Dict[str, Any]:
    """Excel dosyasından form verisini okuyup iç sözlük olarak döndür."""

    filename = get_excel_filename(form_no, base_path=base_path)
    if not os.path.exists(filename):
        raise FormServiceError(f"Form {form_no} bulunamadı. Dosya: {filename}")

    try:
        workbook = openpyxl.load_workbook(filename)
        worksheet = workbook.active
    except Exception as exc:  # pragma: no cover - openpyxl özel hataları
        raise FormServiceError(f"Form {form_no} okunamadı: {exc}") from exc

    raw_data: Dict[str, Any] = {}
    for key_cell, value_cell in worksheet.iter_rows(min_row=2, max_col=2, values_only=True):
        if key_cell:
            raw_data[str(key_cell).strip()] = value_cell

    def parse_datetime_cell(value: Any) -> Tuple[str, str]:
        tarih, saat = "", ""
        if isinstance(value, datetime):
            tarih = value.strftime("%d.%m.%Y")
            saat = value.strftime("%H:%M")
        elif isinstance(value, str):
            cleaned = value.strip()
            if cleaned:
                parts = cleaned.split()
                if len(parts) >= 2:
                    tarih = parts[0]
                    saat = parts[1]
                elif ":" in cleaned:
                    saat = cleaned
                else:
                    tarih = cleaned
        return tarih, saat

    def clean_mola_value(value: Any) -> str:
        if isinstance(value, (int, float)):
            return str(int(value))
        if isinstance(value, str):
            return value.replace("dakika", "").strip()
        return ""

    form_data: Dict[str, Any] = {
        "form_no": form_no,
        "tarih": raw_data.get("Tarih", "") or "",
        "dok_no": raw_data.get("DOK.NO", "") or "",
        "rev_no": raw_data.get("REV.NO/TRH", "") or "",
        "avans": raw_data.get("Avans Tutarı", "") or "",
        "taseron": raw_data.get("Taşeron Şirket", "") or "",
        "gorev_tanimi": raw_data.get("Görevin Tanımı", "") or "",
        "gorev_yeri": raw_data.get("Görev Yeri", "") or "",
        "arac_plaka": raw_data.get("Araç Plaka No", "") or "",
        "hazirlayan": raw_data.get("Hazırlayan", "")
        or raw_data.get("Hazırlayan / Görevlendiren", "")
        or "",
    }

    for index in range(1, 6):
        key = f"Personel {index}"
        form_data[f"personel_{index}"] = raw_data.get(key, "") or ""

    yola_tarih, yola_saat = parse_datetime_cell(raw_data.get("Yola Çıkış"))
    donus_tarih, donus_saat = parse_datetime_cell(raw_data.get("Dönüş"))
    calisma_baslangic_tarih, calisma_baslangic_saat = parse_datetime_cell(
        raw_data.get("Çalışma Başlangıç")
    )
    calisma_bitis_tarih, calisma_bitis_saat = parse_datetime_cell(
        raw_data.get("Çalışma Bitiş")
    )

    form_data.update(
        {
            "yola_cikis_tarih": yola_tarih,
            "yola_cikis_saat": yola_saat,
            "donus_tarih": donus_tarih,
            "donus_saat": donus_saat,
            "calisma_baslangic_tarih": calisma_baslangic_tarih,
            "calisma_baslangic_saat": calisma_baslangic_saat,
            "calisma_bitis_tarih": calisma_bitis_tarih,
            "calisma_bitis_saat": calisma_bitis_saat,
            "mola_suresi": clean_mola_value(raw_data.get("Toplam Mola")) or "",
            "durum": (raw_data.get("DURUM") or "").strip().upper() or "YARIM",
        }
    )

    return form_data


def save_partial_form(form_no: str, form_data: Dict[str, Any], base_path: str = ".") -> Tuple[str, FormStatus]:
    """Formu kısmi olarak kaydet."""

    filename = get_excel_filename(form_no, base_path=base_path)

    try:
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = "Görev Formu"

        header_fill = PatternFill(start_color="FFEB3B", end_color="FFEB3B", fill_type="solid")
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        row = 1
        worksheet[f"A{row}"] = "DELTA PROJE - GÖREV FORMU"
        worksheet[f"A{row}"].font = Font(size=16, bold=True, color="D32F2F")
        worksheet.merge_cells(f"A{row}:B{row}")
        row += 1

        def write_row(label: str, value: Any) -> None:
            nonlocal row
            worksheet[f"A{row}"] = label
            worksheet[f"A{row}"].font = Font(bold=True)
            worksheet[f"A{row}"].fill = header_fill
            worksheet[f"B{row}"] = value
            worksheet[f"A{row}"].border = border
            worksheet[f"B{row}"].border = border
            row += 1

        write_row("Form No", form_no)
        write_row("Tarih", form_data.get("tarih", ""))
        write_row("DOK.NO", form_data.get("dok_no", ""))
        write_row("REV.NO/TRH", form_data.get("rev_no", ""))

        worksheet[f"A{row}"] = "Görevli Personel"
        worksheet[f"A{row}"].font = Font(bold=True)
        worksheet[f"A{row}"].fill = header_fill
        worksheet[f"A{row}"].border = border
        worksheet[f"B{row}"].border = border
        row += 1

        for index in range(5):
            worksheet[f"A{row}"] = f"Personel {index + 1}"
            worksheet[f"A{row}"].border = border
            worksheet[f"B{row}"] = form_data.get(f"personel_{index + 1}", "")
            worksheet[f"B{row}"].border = border
            row += 1

        write_row("Avans Tutarı", form_data.get("avans", ""))
        write_row("Taşeron Şirket", form_data.get("taseron", ""))
        write_row("Görevin Tanımı", form_data.get("gorev_tanimi", ""))
        write_row("Görev Yeri", form_data.get("gorev_yeri", ""))

        worksheet[f"A{row}"] = "DURUM"
        worksheet[f"A{row}"].font = Font(bold=True)
        worksheet[f"A{row}"].fill = PatternFill(start_color="FF9800", end_color="FF9800", fill_type="solid")
        worksheet[f"B{row}"] = "YARIM"
        worksheet[f"B{row}"].fill = PatternFill(start_color="FFC107", end_color="FFC107", fill_type="solid")
        worksheet[f"A{row}"].border = border
        worksheet[f"B{row}"].border = border

        worksheet.column_dimensions["A"].width = 25
        worksheet.column_dimensions["B"].width = 60

        workbook.save(filename)
    except Exception as exc:  # pragma: no cover - dosya sistemi hataları test edilmiyor
        raise FormServiceError(f"Form kaydedilemedi: {exc}") from exc

    return filename, FormStatus(code="YARIM", missing_fields=[])


def save_form(form_no: str, form_data: Dict[str, Any], base_path: str = ".") -> Tuple[str, FormStatus]:
    """Formu tamamlanmış veya yarım olarak kaydet."""

    filename = get_excel_filename(form_no, base_path=base_path)
    status = determine_form_status(form_data)

    try:
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        worksheet.title = "Görev Formu"

        header_fill = PatternFill(start_color="FFEB3B", end_color="FFEB3B", fill_type="solid")
        status_fill = (
            PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid")
            if status.is_complete
            else PatternFill(start_color="FF9800", end_color="FF9800", fill_type="solid")
        )
        status_value_fill = (
            PatternFill(start_color="81C784", end_color="81C784", fill_type="solid")
            if status.is_complete
            else PatternFill(start_color="FFC107", end_color="FFC107", fill_type="solid")
        )

        row = 1
        worksheet[f"A{row}"] = "DELTA PROJE - GÖREV FORMU"
        worksheet[f"A{row}"].font = Font(size=16, bold=True, color="D32F2F")
        worksheet.merge_cells(f"A{row}:B{row}")
        row += 1

        data_map = [
            ("Form No", form_no),
            ("Tarih", form_data.get("tarih", "")),
            ("DOK.NO", form_data.get("dok_no", "")),
            ("REV.NO/TRH", form_data.get("rev_no", "")),
            ("", ""),
            ("Görevli Personel", ""),
        ]

        for label, value in data_map:
            if label:
                worksheet[f"A{row}"] = label
                worksheet[f"A{row}"].font = Font(bold=True)
                worksheet[f"A{row}"].fill = header_fill
                worksheet[f"B{row}"] = value
            row += 1

        for index in range(5):
            worksheet[f"A{row}"] = f"Personel {index + 1}"
            worksheet[f"B{row}"] = form_data.get(f"personel_{index + 1}", "")
            row += 1

        row += 1

        def format_datetime(date_key: str, time_key: str) -> str:
            tarih = (form_data.get(date_key) or "").strip()
            saat = (form_data.get(time_key) or "").strip()
            if tarih and saat:
                return f"{tarih} {saat}"
            return tarih or saat

        mola = (form_data.get("mola_suresi") or "").strip()
        mola_text = f"{mola} dakika" if mola else ""

        all_data = [
            ("Avans Tutarı", form_data.get("avans", "")),
            ("Taşeron Şirket", form_data.get("taseron", "")),
            ("Görevin Tanımı", form_data.get("gorev_tanimi", "")),
            ("Görev Yeri", form_data.get("gorev_yeri", "")),
            ("", ""),
            ("Yola Çıkış", format_datetime("yola_cikis_tarih", "yola_cikis_saat")),
            ("Dönüş", format_datetime("donus_tarih", "donus_saat")),
            (
                "Çalışma Başlangıç",
                format_datetime("calisma_baslangic_tarih", "calisma_baslangic_saat"),
            ),
            (
                "Çalışma Bitiş",
                format_datetime("calisma_bitis_tarih", "calisma_bitis_saat"),
            ),
            ("Toplam Mola", mola_text),
            ("", ""),
            ("Araç Plaka No", form_data.get("arac_plaka", "")),
            ("Hazırlayan", form_data.get("hazirlayan", "")),
            ("", ""),
            ("DURUM", status.code),
        ]

        for label, value in all_data:
            if label:
                worksheet[f"A{row}"] = label
                worksheet[f"A{row}"].font = Font(bold=True)
                if label == "DURUM":
                    worksheet[f"A{row}"].fill = status_fill
                    worksheet[f"B{row}"].fill = status_value_fill
                else:
                    worksheet[f"A{row}"].fill = header_fill
                worksheet[f"B{row}"] = value
            row += 1

        worksheet.column_dimensions["A"].width = 25
        worksheet.column_dimensions["B"].width = 60

        workbook.save(filename)
    except Exception as exc:  # pragma: no cover - dosya sistemi hataları test edilmiyor
        raise FormServiceError(f"Form kaydedilemedi: {exc}") from exc

    return filename, status
