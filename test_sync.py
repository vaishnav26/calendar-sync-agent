import unittest
from unittest.mock import Mock
from googleapiclient.errors import HttpError
from httplib2 import Response
from demo import sample_events
from services import CalendarSyncService, Student, calendar_event_id, event_agent_id


class CalendarBehaviorTests(unittest.TestCase):
    def setUp(self):
        self.student = Student("DEMO001", "Demo Student")
        self.event = sample_events()[0]
        self.sync = CalendarSyncService.__new__(CalendarSyncService)
        self.sync.timezone, self.sync.color_id = "Asia/Kolkata", "10"
        self.sync.service = Mock()

    def test_repeated_event_has_same_id(self):
        self.assertEqual(calendar_event_id(event_agent_id(self.student, self.event)), calendar_event_id(event_agent_id(self.student, sample_events()[0])))

    def test_cancelled_class_has_no_reminder(self):
        body = self.sync._event_body(self.student, sample_events()[1], "demo")
        self.assertTrue(body["summary"].startswith("Cancelled:"))
        self.assertEqual(body["reminders"]["overrides"], [])

    def test_scheduled_event_keeps_timezone_and_reminder(self):
        body = self.sync._event_body(self.student, self.event, "demo")
        self.assertEqual(body["start"]["timeZone"], "Asia/Kolkata")
        self.assertEqual(body["reminders"]["overrides"], [{"method": "popup", "minutes": 10}])

    def test_existing_event_updates_without_insert(self):
        self.sync.service.events.return_value.update.return_value.execute.return_value = {}
        self.assertEqual(self.sync.upsert_events(self.student, [self.event]), {"created": 0, "updated": 1})
        self.sync.service.events.return_value.insert.assert_not_called()

    def test_missing_event_is_inserted_with_deterministic_id(self):
        resource = self.sync.service.events.return_value
        resource.update.return_value.execute.side_effect = HttpError(Response({"status": "404"}), b'{"error":{"message":"Not found"}}')
        resource.insert.return_value.execute.return_value = {}
        self.assertEqual(self.sync.upsert_events(self.student, [self.event]), {"created": 1, "updated": 0})
        self.assertEqual(resource.insert.call_args.kwargs["body"]["id"], calendar_event_id(event_agent_id(self.student, self.event)))


if __name__ == "__main__":
    unittest.main()
