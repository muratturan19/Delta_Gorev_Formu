import sqlite3
from pathlib import Path

import pytest

import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core import form_service


@pytest.fixture
def sample_form_data():
    return {
        "tarih": "01.01.2024",
        "dok_no": "F-001",
        "rev_no": "R1",
        "avans": "1000",
        "taseron": "ABC",
        "gorev_tanimi": "Bakım",
        "gorev_yeri": "İstanbul",
        "gorev_il": "İstanbul",
        "gorev_ilce": "Kadıköy",
        "gorev_firma": "Delta Proje",
        "gorev_tarih": "05.01.2024",
        "personel_1": "Ali",
        "personel_2": "Veli",
        "personel_3": "",
        "personel_4": "",
        "personel_5": "",
        "yola_cikis_tarih": "02.01.2024",
        "yola_cikis_saat": "08:00",
        "calisma_baslangic_tarih": "02.01.2024",
        "calisma_baslangic_saat": "09:00",
        "calisma_bitis_tarih": "02.01.2024",
        "calisma_bitis_saat": "18:00",
        "donus_tarih": "02.01.2024",
        "donus_saat": "19:00",
        "mola_suresi": "30",
        "arac_plaka": "34 ABC 123",
        "hazirlayan": "Ahmet",
        "harcama_bildirimleri": [
            {
                "description": "Yemek",
                "attachments": [
                    {
                        "filename": "fis1.png",
                        "original_name": "Yemek Fişi.png",
                    }
                ],
            },
            {
                "description": "Konaklama",
                "attachments": [],
            },
        ],
    }


def _fetch_form(base_path: Path, form_no: str):
    db_path = base_path / form_service.DB_FILENAME
    assert db_path.exists()
    with sqlite3.connect(db_path) as connection:
        connection.row_factory = sqlite3.Row
        row = connection.execute("SELECT * FROM forms WHERE form_no = ?", (form_no,)).fetchone()
    return row


def test_determine_form_status_complete(sample_form_data):
    status = form_service.determine_form_status(sample_form_data)

    assert status.code == "TAMAMLANDI"
    assert status.is_complete
    assert status.missing_fields == []


def test_determine_form_status_incomplete(sample_form_data):
    sample_form_data["donus_saat"] = ""
    status = form_service.determine_form_status(sample_form_data)

    assert status.code == "YARIM"
    assert not status.is_complete
    assert "donus_saat" in status.missing_fields


def test_get_next_form_no_increments(tmp_path):
    first = form_service.get_next_form_no(base_path=str(tmp_path))
    second = form_service.get_next_form_no(base_path=str(tmp_path))

    assert first == "00001"
    assert second == "00002"
    assert (tmp_path / form_service.DB_FILENAME).exists()


def test_save_partial_form(tmp_path, sample_form_data):
    db_path, status = form_service.save_partial_form(
        "00042", sample_form_data, base_path=str(tmp_path)
    )

    assert status.code == "YARIM"
    assert Path(db_path).exists()
    stored = _fetch_form(tmp_path, "00042")
    assert stored["durum"] == "YARIM"
    assert stored["gorev_yeri"] == sample_form_data["gorev_yeri"]
    assert stored["gorev_il"] == sample_form_data["gorev_il"]
    assert stored["gorev_ilce"] == sample_form_data["gorev_ilce"]
    assert stored["gorev_firma"] == sample_form_data["gorev_firma"]
    assert stored["last_step"] == 0


def test_save_and_load_form(tmp_path, sample_form_data):
    sample_form_data["last_step"] = 4
    db_path, status = form_service.save_form(
        "00077", sample_form_data, base_path=str(tmp_path)
    )

    assert status.code == "TAMAMLANDI"
    assert status.is_complete
    assert Path(db_path).exists()

    stored = _fetch_form(tmp_path, "00077")
    assert stored["durum"] == "TAMAMLANDI"
    assert stored["last_step"] == 4

    loaded = form_service.load_form_data("00077", base_path=str(tmp_path))
    assert loaded["gorev_yeri"] == sample_form_data["gorev_yeri"]
    assert loaded["gorev_il"] == sample_form_data["gorev_il"]
    assert loaded["gorev_ilce"] == sample_form_data["gorev_ilce"]
    assert loaded["gorev_firma"] == sample_form_data["gorev_firma"]
    assert loaded["durum"] == "TAMAMLANDI"
    assert loaded["last_step"] == 4
    assert loaded["gorev_tarih"] == sample_form_data["gorev_tarih"]
    assert loaded["harcama_bildirimleri"] == sample_form_data["harcama_bildirimleri"]


def test_search_forms_filters(tmp_path, sample_form_data):
    form_service.save_form("00001", sample_form_data, base_path=str(tmp_path))

    other = dict(sample_form_data)
    other.update(
        {
            "gorev_yeri": "Ankara",
            "gorev_il": "Ankara",
            "gorev_ilce": "Çankaya",
            "gorev_firma": "XYZ Enerji",
            "personel_1": "Mehmet",
            "personel_2": "Ayşe",
            "yola_cikis_tarih": "15.03.2024",
            "donus_tarih": "16.03.2024",
            "yola_cikis_saat": "07:30",
            "donus_saat": "18:15",
        }
    )
    form_service.save_form("00002", other, base_path=str(tmp_path))

    results = form_service.search_forms(
        person="ali",
        location="istanbul",
        start_date="2024-01-01",
        end_date="2024-01-31",
        base_path=str(tmp_path),
    )
    assert len(results) == 1
    assert results[0]["form_no"] == "00001"

    ankara_results = form_service.search_forms(
        location="Ankara",
        start_date="2024-03-01",
        end_date="2024-03-31",
        base_path=str(tmp_path),
    )
    assert len(ankara_results) == 1
    assert ankara_results[0]["form_no"] == "00002"
    assert "Mehmet" in ankara_results[0]["personel"]


def test_get_reporting_summary(tmp_path, sample_form_data):
    form_service.save_form("00001", sample_form_data, base_path=str(tmp_path))

    second = dict(sample_form_data)
    second.update(
        {
            "gorev_yeri": "Ankara",
            "gorev_il": "Ankara",
            "gorev_ilce": "Çankaya",
            "gorev_firma": "XYZ Enerji",
            "personel_1": "Mehmet",
            "personel_2": "Ayşe",
            "personel_3": "Ali",
            "yola_cikis_tarih": "15.03.2024",
            "yola_cikis_saat": "07:30",
            "donus_tarih": "16.03.2024",
            "donus_saat": "18:15",
            "calisma_baslangic_tarih": "15.03.2024",
            "calisma_baslangic_saat": "09:00",
            "calisma_bitis_tarih": "16.03.2024",
            "calisma_bitis_saat": "16:30",
            "harcama_bildirimleri": [
                {"description": "Yakıt", "attachments": []},
                {"description": "Konaklama", "attachments": []},
            ],
        }
    )
    form_service.save_form("00002", second, base_path=str(tmp_path))

    summary = form_service.get_reporting_summary(
        start_date="2024-01-01",
        end_date="2024-12-31",
        base_path=str(tmp_path),
    )

    assert summary["total_forms"] == 2
    assert summary["unique_person_count"] >= 3
    assert summary["travel_hours"]["total"] > 0
    assert summary["work_hours"]["total"] > 0
    assert len(summary["expense_chart"]["labels"]) == 2
    assert any("Ankara" in item["label"] for item in summary["locations"])
