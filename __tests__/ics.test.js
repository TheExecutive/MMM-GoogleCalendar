jest.mock("logger", () => ({ log: jest.fn(), info: jest.fn(), debug: jest.fn(), warn: jest.fn(), error: jest.fn() }), { virtual: true });
const { fetchIcs, toGoogleItems } = require("../ics.js");

const ICS = `BEGIN:VCALENDAR
VERSION:2.0
BEGIN:VTIMEZONE
TZID:Eastern Standard Time
BEGIN:STANDARD
DTSTART:16010101T020000
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
RRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=11
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:16010101T020000
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
RRULE:FREQ=YEARLY;BYDAY=2SU;BYMONTH=3
END:DAYLIGHT
END:VTIMEZONE
BEGIN:VEVENT
UID:standup
SUMMARY:IS Team Standup
DTSTART;TZID=Eastern Standard Time:20261123T093000
DTEND;TZID=Eastern Standard Time:20261123T094500
RRULE:FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR
EXDATE;TZID=Eastern Standard Time:20261127T093000
LOCATION:Microsoft Teams Meeting
END:VEVENT
BEGIN:VEVENT
UID:off
SUMMARY:Day off
DTSTART;VALUE=DATE:20261130
DTEND;VALUE=DATE:20261201
END:VEVENT
END:VCALENDAR`;

beforeEach(() => {
  jest.useFakeTimers({ now: new Date("2026-11-26T16:00:00Z"), doNotFake: ["nextTick", "setImmediate", "queueMicrotask"] });
  global.fetch = jest.fn(async () => ({ ok: true, text: async () => ICS }));
});
afterEach(() => jest.useRealTimers());

test("expands Outlook recurrences with Windows time zones and exceptions, in Google's shape", async () => {
  const items = await fetchIcs("webcal://example.com/cal.ics", { maximumEntries: 50, maximumNumberOfDays: 7 });
  expect(global.fetch.mock.calls[0][0]).toBe("https://example.com/cal.ics");
  const standups = items.filter(i => i.summary === "IS Team Standup").map(i => i.start.dateTime);
  expect(standups[0]).toBe("2026-11-30T09:30:00-05:00"); // 11 AM Thu 26: today's has passed; Fri 27 is an exception
  expect(standups.some(s => s.startsWith("2026-11-27"))).toBe(false);
  expect(items.find(i => i.summary === "Day off")).toMatchObject({ start: { date: "2026-11-30" }, end: { date: "2026-12-01" }, status: "confirmed" });
});

test("respects maximumEntries and reports HTTP failures", async () => {
  expect((await fetchIcs("https://x/c.ics", { maximumEntries: 2, maximumNumberOfDays: 30 })).length).toBe(2);
  global.fetch = jest.fn(async () => ({ ok: false, status: 404 }));
  await expect(fetchIcs("https://x/c.ics", {})).rejects.toThrow("404");
});

test("private events are marked private", () => {
  expect(toGoogleItems([{ title: "x", startDate: "0", endDate: "1", fullDayEvent: false, class: "PRIVATE" }], 5)[0].visibility).toBe("private");
});
