import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core import task_request_service, user_service  # noqa: E402


@pytest.fixture
def base_path(tmp_path):
    return tmp_path


@pytest.fixture
def requester(base_path):
    user_service.ensure_default_users(base_path=str(base_path))
    admins = user_service.list_users_by_role("admin", base_path=str(base_path))
    assert admins
    return admins[0]


def test_create_and_list_task_request(base_path, requester):
    created = task_request_service.create_task_request(
        customer_name="ABC Şirketi",
        customer_phone="0532 123 45 67",
        customer_email="iletisim@abc.com",
        customer_address="Ataşehir Plaza Kat:5",
        request_description="Klima arızası bildirildi, sistem soğutmuyor.",
        requirements="2 teknisyen, R410A gazı",
        urgency="normal",
        requested_by_user_id=requester.id,
        base_path=str(base_path),
    )

    assert created["customer_name"] == "ABC Şirketi"
    requests = task_request_service.list_task_requests(base_path=str(base_path))
    assert len(requests) == 1
    assert requests[0]["display_id"] == "#001"
    assert requests[0]["urgency_label"] == "Normal"
    assert "Klima arızası" in requests[0]["request_summary"]


def test_update_status_and_notes(base_path, requester):
    created = task_request_service.create_task_request(
        customer_name="XYZ Enerji",
        customer_phone=None,
        customer_email=None,
        customer_address="",
        request_description="Trafo bakımı için keşif talebi.",
        requirements=None,
        urgency="urgent",
        requested_by_user_id=requester.id,
        base_path=str(base_path),
    )

    updated = task_request_service.update_task_request_status(
        created["id"], status="in_progress", base_path=str(base_path)
    )
    assert updated["status"] == "in_progress"
    note_updated = task_request_service.update_task_request_notes(
        created["id"], notes="Ekip yönlendirildi.", base_path=str(base_path)
    )
    assert note_updated["notes"] == "Ekip yönlendirildi."


def test_mark_converted_updates_status(base_path, requester):
    created = task_request_service.create_task_request(
        customer_name="Delta Endüstri",
        customer_phone="0212 000 00 00",
        customer_email=None,
        customer_address="İkitelli OSB",
        request_description="Üretim hattında arıza var, hızlı müdahale gerekiyor.",
        requirements="Kompresör testi ekipmanı",
        urgency="very_urgent",
        requested_by_user_id=requester.id,
        base_path=str(base_path),
    )

    converted = task_request_service.mark_converted(
        created["id"], form_no="00010", base_path=str(base_path)
    )
    assert converted["status"] == "converted"
    assert converted["converted_form_no"] == "00010"
