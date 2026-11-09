from __future__ import annotations

from datetime import datetime
from zoneinfo import ZoneInfo

import pytest

from outlook_draft import availability as av

LONDON = ZoneInfo("Europe/London")
START = datetime(2026, 11, 10, 9, 0, tzinfo=LONDON)

# Real AvailabilityView strings from 2026-11-09, window Tue 10 Nov 09:00 to Wed 18 Nov 17:30, 30-minute steps.
ROSS = "11111111111221110000000000000000000000000000000011111111112221110000000000000000000000000000000011111111111221110000000000000000000000000000000010000000000220000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000022222222220220000000000000000000000000000000000000100000011221221000000000000000000000000000000000000010010220000"
CELINE = "22222111222222222222200000000000000000000000002222222222222222122222200000000000000000000000002222222222222212222222220000000000000000000000222222222222222222222222000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000200022221222222221222222200000000000000000000000222022222222222222222222200000003333333333333333333333333333333333333"
PHILIP = "22220021002000000000000000000000000000000000000222000020112200000000000000000000000000000000000222200021002222000000000000000000000000000000000222200022200000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000222000020200000000000000000000000000000000000000222220021002000000000000000000000000000000000000222000220002200022"
TARIK = "20000001222222222200000000000000000000000000022222222222222222222200000000000000000000000000000001220000002222222200000000000000000000000000002022000000222222222200000000000000000000000000000000000000000000222200000000000000000000000000000000000000000000222200000000000000000000000000022222222222221120222200000000000000000000000000222222222222222222222200000000000000000000000000000022000020020022222"

PHILIP_HOURS = {
    "DaysOfWeek": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"],
    "StartTime": "08:45:00.0000000",
    "EndTime": "17:00:00.0000000",
    "TimeZone": {"Name": "Romance Standard Time"},
}
CELINE_HOURS = {
    "DaysOfWeek": ["Monday", "Tuesday", "Wednesday"],
    "StartTime": "00:00:00.0000000",
    "EndTime": "17:00:00.0000000",
    "TimeZone": {"Name": "Singapore Standard Time"},
}


def _items(with_hours: bool = False) -> list[dict]:
    views = {"ross": ROSS, "celine": CELINE, "philip": PHILIP, "tarik": TARIK}
    hours = {"celine": CELINE_HOURS, "philip": PHILIP_HOURS} if with_hours else {}
    return [
        {"ScheduleId": name, "AvailabilityView": view, "WorkingHours": hours.get(name)}
        for name, view in views.items()
    ]


def test_decode_ranges_merges_runs_and_skips_free_time():
    ranges = av.decode_ranges(PHILIP[:12], START, 30)

    assert [(r["status"], r["start"], r["end"]) for r in ranges] == [
        ("busy", "2026-11-10T09:00", "2026-11-10T11:00"),
        ("busy", "2026-11-10T12:00", "2026-11-10T12:30"),
        ("tentative", "2026-11-10T12:30", "2026-11-10T13:00"),
        ("busy", "2026-11-10T14:00", "2026-11-10T14:30"),
    ]
    assert ranges[0]["label"] == "Tue 10 Nov 09:00-11:00"


def test_decode_ranges_names_out_of_office_and_unknown_codes():
    ranges = av.decode_ranges("39", START, 60)

    assert [r["status"] for r in ranges] == ["outOfOffice", "unknown(9)"]


def test_common_free_matches_the_hand_decoded_slots():
    report = av.availability_report(_items(), START, 30, av.make_options())

    labels = {w["label"]: w["tentative"] for w in report["commonFree"]}
    assert labels["Tue 10 Nov 11:30-12:00"] == ["ross", "celine"]
    assert labels["Mon 16 Nov 15:30-16:00"] == ["celine"]
    assert all(not w["clear"] for w in report["commonFree"])
    assert "Wed 18 Nov" not in " ".join(labels)  # Celine is out of office


def test_tentative_busy_leaves_no_common_slot_in_that_window():
    options = av.make_options(tentative_is_free=False)

    assert av.availability_report(_items(), START, 30, options)["commonFree"] == []


def test_find_slots_ranks_fewest_issues_first():
    report = av.find_slots_report(_items(), START, 30, 30, 3, av.make_options())

    assert report["slots"][0]["label"] == "Mon 16 Nov 15:30-16:00"
    assert report["slots"][0]["tentative"] == ["celine"]
    assert len(report["slots"]) == 3


def test_working_hours_flags_people_outside_their_own_hours():
    report = av.find_slots_report(_items(with_hours=True), START, 30, 30, 50, av.make_options())

    best = next(s for s in report["slots"] if s["label"] == "Mon 16 Nov 15:30-16:00")
    # 15:30 London is 23:30 in Singapore; Philip's 08:45-17:00 Paris covers 16:30-17:00 Paris.
    assert best["outsideWorkingHours"] == ["celine"]


def test_parse_working_hours_maps_windows_zone_names():
    hours = av.parse_working_hours(PHILIP_HOURS)

    assert hours is not None
    assert hours.label == "Mon-Fri 08:45-17:00 Romance Standard Time"
    late = datetime(2026, 11, 10, 16, 30, tzinfo=LONDON)  # 17:30 in Paris
    assert av.in_working_hours(hours, late.replace(hour=15), late.replace(hour=15, minute=30))
    assert not av.in_working_hours(hours, late, late.replace(hour=17))


def test_unknown_working_hours_zone_gives_none():
    assert av.parse_working_hours({**PHILIP_HOURS, "TimeZone": {"Name": "Nowhere Time"}}) is None


def test_errors_are_reported_and_left_out_of_slots():
    items = [*_items(), {"ScheduleId": "room@example.com", "Error": {"Message": "No access"}}]

    report = av.availability_report(items, START, 30, av.make_options())

    assert report["unknown"] == [{"schedule": "room@example.com", "error": "No access"}]
    assert "room@example.com" in report["warning"]
    assert report["commonFree"]


def test_short_or_unknown_views_count_as_busy():
    items = [
        {"ScheduleId": "a", "AvailabilityView": "0000"},
        {"ScheduleId": "b", "AvailabilityView": "09"},
    ]

    slots = av.candidate_slots(av.schedules_from_response(items)[0], START, 30, 30, av.make_options())

    assert [av.format_local(s.start) for s in slots] == ["2026-11-10T09:00"]


def test_merge_windows_joins_back_to_back_slots_and_splits_on_new_flags():
    items = [{"ScheduleId": "a", "AvailabilityView": "0001"}]

    report = av.availability_report(items, START, 30, av.make_options())

    assert [(w["label"], w["minutes"], w["tentative"]) for w in report["commonFree"]] == [
        ("Tue 10 Nov 09:00-10:30", 90, []),
        ("Tue 10 Nov 10:30-11:00", 30, ["a"]),
    ]


def test_slot_time_steps_in_real_time_across_dst_change():
    start = datetime(2026, 10, 24, 23, 0, tzinfo=LONDON)  # clocks go back at 02:00 BST on 25 Oct

    assert av.slot_time(start, 60, 3).strftime("%H:%M %Z") == "01:00 GMT"


def test_day_window_and_weekend_filters():
    options = av.make_options(day_start="09:00", day_end="09:30", weekdays_only=True)
    saturday = datetime(2026, 11, 14, 9, 0, tzinfo=LONDON)
    items = [{"ScheduleId": "a", "AvailabilityView": "0" * 4}]

    assert len(av.availability_report(items, START, 30, options)["commonFree"]) == 1
    assert av.availability_report(items, saturday, 30, options)["commonFree"] == []


@pytest.mark.parametrize("bad", [{"day_start": "18:00", "day_end": "09:00"}, {"day_start": "nine"}])
def test_make_options_rejects_bad_times(bad):
    with pytest.raises(ValueError):
        av.make_options(**bad)


def test_interval_helpers():
    assert av.interval_for_duration(45) == 15
    assert av.interval_for_duration(60) == 30
    with pytest.raises(ValueError):
        av.interval_for_duration(0)
    with pytest.raises(ValueError):
        av.check_interval(2)
