"""Offline walkthrough with fictional data; no OAuth or network calls."""
import argparse
import json
from dataclasses import asdict
from pathlib import Path
from openpyxl import Workbook
from services import CalendarSyncService, CourseEnrollment, Student, TimetableEvent, calendar_event_id, event_agent_id


def sample_events():
    return [
        TimetableEvent("2027-01-11", "09:00", "10:30", "26MBA401", "Digital Transformation", "A", location="Demo Room 1", source_sheet="Synthetic demo"),
        TimetableEvent("2027-01-12", "14:00", "15:30", "26MBA402", "Product Analytics", "A", source_sheet="Synthetic demo", cancelled=True),
    ]


def write_sample(path=Path("sample_sections.xlsx")):
    wb = Workbook()
    ws = wb.active
    ws.title = "Synthetic enrollments"
    ws.append(["Section", "CourseCode", "StudentId", "Name"])
    ws.append(["26MBA401: Digital Transformation [A]", "26MBA401", "DEMO001", "Demo Student"])
    ws.append(["26MBA402: Product Analytics [A]", "26MBA402", "DEMO001", "Demo Student"])
    wb.save(path)
    return path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write-sample", action="store_true")
    args = parser.parse_args()
    if args.write_sample:
        print(f"Created fictional enrollment workbook: {write_sample()}")
        return
    student = Student("DEMO001", "Demo Student", [CourseEnrollment("26MBA401", "Digital Transformation", "A"), CourseEnrollment("26MBA402", "Product Analytics", "A")])
    formatter = CalendarSyncService.__new__(CalendarSyncService)
    formatter.timezone, formatter.color_id = "Asia/Kolkata", "10"
    events = sample_events()
    payloads = []
    for event in events:
        identifier = event_agent_id(student, event)
        payloads.append({"id": calendar_event_id(identifier), **formatter._event_body(student, event, identifier)})
    print(json.dumps({"mode": "offline; nothing sent to Google", "student": asdict(student), "events": payloads}, indent=2))
    print("\nScheduled class: 10-minute reminder. Cancelled class: labelled cancelled, no reminder.")
    print("The same event inputs produce the same IDs on every run.")


if __name__ == "__main__":
    main()
