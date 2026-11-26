/* Read an ICS feed (e.g. a published Outlook calendar) and return events shaped like
   Google Calendar API items, so the rest of the module treats them the same way.
   Recurrence, exceptions and Windows time-zone names are handled by MagicMirror's
   own calendar parser. */
const path = require("path");
const moment = require("moment-timezone");
const ical = require("node-ical");
const CalendarFetcherUtils = require(path.join(__dirname, "../../defaultmodules/calendar/calendarfetcherutils.js"));

function toGoogleItems(events, maxResults) {
  return events.slice(0, maxResults).map((e, i) => {
    const start = moment(parseInt(e.startDate, 10)), end = moment(parseInt(e.endDate, 10));
    const when = m => (e.fullDayEvent ? { date: m.format("YYYY-MM-DD") } : { dateTime: m.format() });
    return {
      id: `ics-${e.startDate}-${i}`,
      status: "confirmed",
      summary: e.title,
      location: e.location || undefined,
      start: when(start),
      end: when(end),
      visibility: e.class === "PRIVATE" ? "private" : undefined
    };
  });
}

async function fetchIcs(url, { maximumEntries, maximumNumberOfDays = 365, excludedEvents = [] } = {}) {
  const res = await fetch(url.replace(/^webcal:/i, "https:"), { signal: AbortSignal.timeout(30000) });
  if (!res.ok) throw new Error(`ICS HTTP ${res.status}`);
  const config = { maximumNumberOfDays, excludedEvents, includePastEvents: false };
  let raw = await res.text();
  try {
    raw = await CalendarFetcherUtils.preFilterICS(raw, config); // speed-up only
  } catch {
    // keep the full feed
  }
  // filterEvents checks a series' own title; changed single meetings (e.g. Outlook's
  // "Canceled: ...") carry their own title, so check every meeting again.
  const events = CalendarFetcherUtils.filterEvents(ical.sync.parseICS(raw), config)
    .filter(e => !CalendarFetcherUtils.shouldEventBeExcluded(config, e.title).excluded);
  return toGoogleItems(events, maximumEntries);
}

module.exports = { fetchIcs, toGoogleItems };
