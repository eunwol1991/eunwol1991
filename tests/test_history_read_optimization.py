class _Cell:
    def __init__(self, value):
        self.value = value


class _FakeWS:
    def __init__(self):
        self.max_row = 100000
        self.calls = 0
        self._data = {
            (4, 1): "19/02/2026",
            (4, 2): "AEON",
            (4, 3): "Tampines",
            (4, 4): "Chicken Nugget",
            (5, 1): "02/01/2024",
            (5, 2): "AEON",
            (5, 3): "Tampines Mall",
            (5, 4): "Chicken Nugget",
            (90000, 1): "01/01/2020",
            (90000, 2): "OLD",
            (90000, 3): "OLD",
            (90000, 4): "OLD",
        }

    def cell(self, row, column):
        self.calls += 1
        return _Cell(self._data.get((row, column)))


def test_build_history_records_stops_after_empty_streak():
    from eunwol1991.projects.function.delivery_assistant.history_index import (
        build_history_records,
    )

    ws = _FakeWS()
    records = build_history_records(
        ws,
        header_row=3,
        columns={"date": 1, "customer": 2, "outlet": 3, "description": 4},
        max_empty_streak=50,
    )

    assert len(records) == 2
    assert records[0]["row_idx"] == 4
    assert records[1]["row_idx"] == 5
    assert ws.calls < 600
