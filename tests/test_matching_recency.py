from datetime import date


def test_recency_bias_prefers_newer_row_when_similarity_close():
    from eunwol1991.projects.function.delivery_assistant.matching import rank_candidates

    query = {
        "customer": "aeon",
        "outlet": "tampines",
        "description": "chicken nugget",
    }
    records = [
        {
            "row_idx": 20,
            "customer": "aeon",
            "outlet": "tampines mall",
            "description": "chicken nugget",
            "record_date": date(2024, 1, 2),
        },
        {
            "row_idx": 7,
            "customer": "aeon",
            "outlet": "tampines",
            "description": "chicken nugget",
            "record_date": date(2026, 2, 19),
        },
    ]

    ranked = rank_candidates(query, records, limit=2)
    assert ranked[0]["record"]["row_idx"] == 7


def test_confidence_thresholds():
    from eunwol1991.projects.function.delivery_assistant.matching import (
        confidence_level,
    )

    assert confidence_level(95.0, 83.0) == "high"
    assert confidence_level(85.0, 80.0) == "medium"
    assert confidence_level(77.0, 76.0) == "low"
    assert confidence_level(60.0, 55.0) == "none"
