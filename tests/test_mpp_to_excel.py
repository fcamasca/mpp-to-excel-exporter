from datetime import datetime

from mpp_to_excel import extract_rows, format_duration, parse_date


def test_format_duration_uses_eight_hour_workdays():
    assert format_duration(0) == "0 días"
    assert format_duration(30 * 60) == "30 min"
    assert format_duration(60 * 60) == "1 hora"
    assert format_duration(8 * 60 * 60) == "1 día"
    assert format_duration(12 * 60 * 60) == "1.5 días"


def test_parse_date_accepts_iso_and_invalid_values():
    assert parse_date("2026-02-06T08:00:00") == datetime(2026, 2, 6, 8)
    assert parse_date("fecha inválida") is None
    assert parse_date(None) is None


def test_extract_rows_preserves_hierarchy_and_resources():
    data = {
        "resources": [{"unique_id": 7, "name": "Analista"}],
        "assignments": [{"task_unique_id": 2, "resource_unique_id": 7}],
        "tasks": [
            {"unique_id": 1, "name": "Proyecto", "summary": True, "outline_level": 1},
            {
                "unique_id": 2,
                "name": "Implementación",
                "outline_level": 2,
                "duration": 28800,
                "work": 14400,
                "percent_complete": 50,
            },
        ],
    }

    rows, summary_flags = extract_rows(data)

    assert len(rows) == 2
    assert rows[1]["Nombre de tarea"] == "      Implementación"
    assert rows[1]["Nombres de los recursos"] == "Analista"
    assert rows[1]["% completado"] == 0.5
    assert summary_flags == [True, False]
