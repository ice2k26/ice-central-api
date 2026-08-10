/** One-off patch: colour-code the ICE 2026 program calendar (15–17 Aug 2026)
 *  by activity category, so the printed-schedule colour legend is reflected in
 *  Google Calendar (and readable back via each event's colorId in the feed).
 *
 *  Run colorIce2026Program() once from the Apps Script editor. Idempotent and
 *  safe to re-run — it just re-asserts each event's colour. Categorisation is
 *  by title keyword (not event id), so it survives renames and the duplicate
 *  titles ("Dinner", "Summarize the Day's Activities"). Requires the writable
 *  calendar scope (https://www.googleapis.com/auth/calendar) in appsscript.json.
 *
 *  Legend → Google event colour (colorId):
 *    Collective Activity    ORANGE   (6)   Interactive Discussion PALE_GREEN (2)
 *    Hands on Activity      CYAN     (7)   Refreshment            GRAY       (8)
 *    Invited Talk           YELLOW   (5)   Team Presentation      PALE_RED   (4)
 *    Finale Event           MAUVE    (3)
 */
function colorIce2026Program() {
  var calId =
    getConfig_('PROGRAM_CALENDAR_ID_ice2026', '') ||
    getConfig_('PROGRAM_CALENDAR_ID', '') ||
    'c_77ef808e169d66dc1b79a4ba4c3e0dbb0e51fdfc2abefacb61e33f1a6a6f1e84@group.calendar.google.com';
  var cal = CalendarApp.getCalendarById(calId);
  if (!cal) throw new Error('Calendar not accessible: ' + calId);
  var tz = cal.getTimeZone() || Session.getScriptTimeZone();

  var C = CalendarApp.EventColor;
  var COLOR = {
    collective:   C.ORANGE,     // 6  Tangerine
    discussion:   C.PALE_GREEN,  // 2  Sage
    handson:      C.CYAN,        // 7  Peacock
    refreshment:  C.GRAY,        // 8  Graphite
    talk:         C.YELLOW,      // 5  Banana
    presentation: C.PALE_RED,    // 4  Flamingo
    finale:       C.MAUVE,       // 3  Grape
  };

  var start = Utilities.parseDate('2026-08-15 00:00', tz, 'yyyy-MM-dd HH:mm');
  var end   = Utilities.parseDate('2026-08-18 00:00', tz, 'yyyy-MM-dd HH:mm');
  var events = cal.getEvents(start, end);

  var counts = {};
  events.forEach(function (ev) {
    var cat = categorizeIce2026_(ev.getTitle());
    ev.setColor(COLOR[cat]);
    counts[cat] = (counts[cat] || 0) + 1;
  });

  Logger.log('ICE 2026 colour patch: %s events on %s [%s] — %s',
    events.length, calId, tz, JSON.stringify(counts));
}

/** Map a program event title to one of the seven legend categories.
 *  Order matters: more specific rules first, Hands-on Activity is the default. */
function categorizeIce2026_(title) {
  var t = String(title).toLowerCase();

  // Finale Event — closing block
  if (/finale|vip|certificate|clean up|debrief/.test(t)) return 'finale';

  // Refreshment — breaks & meals. \btea\b so it doesn't match inside "team".
  if (/\btea\b|lunch|dinner|refreshment/.test(t)) return 'refreshment';

  // Invited Talk
  if (/invited talk/.test(t)) return 'talk';

  // Team Presentation (covers "Team Presentation(s)" + the finalize-prototypes pitch)
  if (/team presentation|finalize prototypes/.test(t)) return 'presentation';

  // Collective Activity — plenary welcomes
  if (/introduction & welcome|welcome to day/.test(t)) return 'collective';

  // Interactive Discussion — briefings / share-outs / summaries
  if (/workshop objectives|ai sneak|summarize the day|what do prototypes prototype|important aspects/.test(t))
    return 'discussion';

  // Default: Hands on Activity (challenges, defining, prototyping, iterating, team meetings)
  return 'handson';
}
