"""Comprehensive test suite for the Analysis & Audit functions."""
import pytest
from coordinator_logic import calculate_slots_and_dues, detect_strict_conflicts


def make_row(group_code="G01", center_code="C001", room_code="R01", area="Area1", city="Cairo",
             center_name="Center A", day="Saturday", time="10:00 AM",
             main_coordinator="Ahmed", main_phone="+201012345678",
             backup_coordinator="Sara", backup_phone="+201098765432"):
    return {
        "group_code": group_code, "center_code": center_code, "room_code": room_code,
        "area": area, "city": city, "center_name": center_name, "day": day, "time": time,
        "main_coordinator": main_coordinator, "main_phone": main_phone,
        "backup_coordinator": backup_coordinator, "backup_phone": backup_phone,
    }


# ---------------------------------------------------------------------------
# Slot & Dues Tests
# ---------------------------------------------------------------------------

def test_same_city_different_centers():
    """
    Same city but different center_code → two separate slots, NOT merged.
    Also triggers a coordinator conflict (same person, same day/time, >1 center).
    """
    rows = [
        make_row(center_code="C001", center_name="Center 1", city="Cairo",
                 day="Saturday", time="10:00 AM", main_coordinator="Ahmed"),
        make_row(center_code="C002", center_name="Center 2", city="Cairo",
                 day="Saturday", time="10:00 AM", main_coordinator="Ahmed"),
    ]
    analysis, totals = calculate_slots_and_dues(rows)
    ahmed = next((x for x in analysis if x["Coordinator"] == "Ahmed"), None)
    assert ahmed is not None, "Ahmed should appear in analysis"
    assert ahmed["Main Slots"] == 2, "Different center codes = different slots"

    conflicts = detect_strict_conflicts(rows)
    assert len(conflicts["coordinator_conflicts"]) > 0, (
        "Should flag coordinator conflict for same day/time at different centers"
    )


def test_multiple_groups_same_slot():
    """3 groups in the same Center/Day/Time → exactly 1 slot = 250 EGP."""
    rows = [
        make_row(center_code="C001", day="Saturday", time="10:00 AM",
                 group_code="G01", main_coordinator="Ahmed"),
        make_row(center_code="C001", day="Saturday", time="10:00 AM",
                 group_code="G02", main_coordinator="Ahmed"),
        make_row(center_code="C001", day="Saturday", time="10:00 AM",
                 group_code="G03", main_coordinator="Ahmed"),
    ]
    analysis, totals = calculate_slots_and_dues(rows, slot_price=250)
    ahmed = next((x for x in analysis if x["Coordinator"] == "Ahmed"), None)
    assert ahmed is not None
    assert ahmed["Main Slots"] == 1, "Same center/day/time = 1 slot regardless of group count"
    assert ahmed["Amount (EGP)"] == 250


def test_coordinator_double_booked():
    """Coordinator in Center A and Center B at the exact same Day/Time → conflict."""
    rows = [
        make_row(center_code="C001", day="Saturday", time="10:00 AM",
                 main_coordinator="Ahmed"),
        make_row(center_code="C002", day="Saturday", time="10:00 AM",
                 main_coordinator="Ahmed"),
    ]
    conflicts = detect_strict_conflicts(rows)
    coord_conflicts = conflicts["coordinator_conflicts"]
    assert len(coord_conflicts) > 0, "Should detect double booking"
    centers = coord_conflicts[0]["Centers"]
    assert "C001" in centers and "C002" in centers, "Both centers must be listed"


def test_empty_and_zero_coordinators():
    """Rows with '0', '0.0', 'None', 'NaN', '<NA>', '' → all completely ignored."""
    bad_names = ["0", "0.0", "None", "NaN", "<NA>", ""]
    rows = [
        make_row(center_code=f"C{i:03d}", main_coordinator=name, backup_coordinator=name)
        for i, name in enumerate(bad_names)
    ]
    analysis, totals = calculate_slots_and_dues(rows)
    assert len(analysis) == 0, "All null names should be filtered out"
    assert totals["Total Slots"] == 0

    conflicts = detect_strict_conflicts(rows)
    assert len(conflicts["coordinator_conflicts"]) == 0, "No conflicts for null names"
    assert len(conflicts["room_conflicts"]) == 0


def test_room_conflict_detected():
    """Same room in same center with 2 different groups at same day/time → room conflict."""
    rows = [
        make_row(center_code="C001", room_code="R01", day="Saturday",
                 time="10:00 AM", group_code="G01"),
        make_row(center_code="C001", room_code="R01", day="Saturday",
                 time="10:00 AM", group_code="G02"),
    ]
    conflicts = detect_strict_conflicts(rows)
    room_conflicts = conflicts["room_conflicts"]
    assert len(room_conflicts) > 0, "Room conflict should be detected"
    groups = room_conflicts[0]["Groups"]
    assert "G01" in groups and "G02" in groups


def test_no_conflicts_clean_data():
    """Clean data with completely separate assignments → no conflicts."""
    rows = [
        make_row(center_code="C001", day="Saturday", time="10:00 AM",
                 main_coordinator="Ahmed", room_code="R01", group_code="G01"),
        make_row(center_code="C002", day="Sunday", time="11:00 AM",
                 main_coordinator="Mohamed", room_code="R02", group_code="G02"),
        make_row(center_code="C003", day="Monday", time="12:00 PM",
                 main_coordinator="Sara", room_code="R03", group_code="G03"),
    ]
    conflicts = detect_strict_conflicts(rows)
    assert len(conflicts["coordinator_conflicts"]) == 0
    assert len(conflicts["room_conflicts"]) == 0


def test_backup_coordinator_slots():
    """Sara as backup in 2 different centers on different days → 2 backup slots, 500 EGP."""
    rows = [
        make_row(center_code="C001", day="Saturday", time="10:00 AM",
                 main_coordinator="Ahmed", backup_coordinator="Sara"),
        make_row(center_code="C002", day="Sunday", time="11:00 AM",
                 main_coordinator="Mohamed", backup_coordinator="Sara"),
    ]
    analysis, _ = calculate_slots_and_dues(rows, slot_price=250)
    sara = next((x for x in analysis if x["Coordinator"] == "Sara"), None)
    assert sara is not None, "Sara should appear in analysis as backup"
    assert sara["Main Slots"] == 0
    assert sara["Backup Slots"] == 2
    assert sara["Total Slots"] == 2
    assert sara["Amount (EGP)"] == 500


def test_whitespace_normalization():
    """' Ahmed ' and 'Ahmed' are the same person → 1 slot, not 2."""
    rows = [
        make_row(center_code="C001", day="Saturday", time="10:00 AM",
                 group_code="G01", main_coordinator=" Ahmed "),
        make_row(center_code="C001", day="Saturday", time="10:00 AM",
                 group_code="G02", main_coordinator="Ahmed"),
    ]
    analysis, _ = calculate_slots_and_dues(rows)
    ahmeds = [x for x in analysis if "Ahmed" in x["Coordinator"]]
    assert len(ahmeds) == 1, "Whitespace variants should merge into one person"
    assert ahmeds[0]["Main Slots"] == 1, "Same slot = 1 slot"


def test_grand_total_calculation():
    """Ahmed: 2 slots, Sara: 3 slots, Mohamed: 1 slot → 6 total, 1500 EGP."""
    rows = [
        # Ahmed — 2 slots (different center/day combos)
        make_row(center_code="C001", day="Saturday", time="10:00 AM",
                 main_coordinator="Ahmed", backup_coordinator="None"),
        make_row(center_code="C002", day="Sunday", time="11:00 AM",
                 main_coordinator="Ahmed", backup_coordinator="None"),
        # Sara — 3 slots
        make_row(center_code="C003", day="Monday", time="12:00 PM",
                 main_coordinator="Sara", backup_coordinator="None"),
        make_row(center_code="C004", day="Tuesday", time="1:00 PM",
                 main_coordinator="Sara", backup_coordinator="None"),
        make_row(center_code="C005", day="Wednesday", time="2:00 PM",
                 main_coordinator="Sara", backup_coordinator="None"),
        # Mohamed — 1 slot
        make_row(center_code="C006", day="Thursday", time="3:00 PM",
                 main_coordinator="Mohamed", backup_coordinator="None"),
    ]
    analysis, totals = calculate_slots_and_dues(rows, slot_price=250)
    assert totals["Total Slots"] == 6, f"Expected 6 total slots, got {totals['Total Slots']}"
    assert totals["Total Amount (EGP)"] == 1500, f"Expected 1500 EGP, got {totals['Total Amount (EGP)']}"
