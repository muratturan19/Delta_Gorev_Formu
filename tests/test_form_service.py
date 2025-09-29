import os
import sys
from pathlib import Path
from typing import Optional

import openpyxl
import pytest

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
    }


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
    config_path = tmp_path / form_service.CONFIG_FILE_NAME
    assert config_path.exists()


def _find_status_value(filename: str) -> Optional[str]:
    workbook = openpyxl.load_workbook(filename)
    worksheet = workbook.active
    for cell in worksheet['A']:
        if cell.value == 'DURUM':
            return worksheet[f'B{cell.row}'].value
    return None
 

def test_save_partial_form(tmp_path, sample_form_data):
    filename, status = form_service.save_partial_form(
        "00042", sample_form_data, base_path=str(tmp_path)
    )

    assert status.code == "YARIM"
    assert os.path.exists(filename)
    assert _find_status_value(filename) == "YARIM"


def test_save_and_load_form(tmp_path, sample_form_data):
    filename, status = form_service.save_form(
        "00077", sample_form_data, base_path=str(tmp_path)
    )

    assert status.code == "TAMAMLANDI"
    assert status.is_complete
    assert os.path.exists(filename)
    assert _find_status_value(filename) == "TAMAMLANDI"

    loaded = form_service.load_form_data("00077", base_path=str(tmp_path))
    assert loaded["gorev_yeri"] == sample_form_data["gorev_yeri"]
    assert loaded["durum"] == "TAMAMLANDI"
