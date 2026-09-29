"""Read Victorian government calendars without changing the active roster calendar."""
from datetime import date, datetime, timedelta, timezone
from html.parser import HTMLParser
from urllib.parse import urljoin, urlparse
import copy
import hashlib
import json
import re

from urllib.request import build_opener, HTTPRedirectHandler, Request
from urllib.error import HTTPError

SOURCE_PAGE = "https://www.vic.gov.au/ical"


def stamp():
    return datetime.now(timezone.utc).isoformat()


class CalendarLinks(HTMLParser):
    def __init__(self):
        super().__init__()
        self.links = []
        self.href = None
        self.label = ""

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            self.href = dict(attrs).get("href")
            self.label = ""

    def handle_data(self, data):
        if self.href:
            self.label += data

    def handle_endtag(self, tag):
        if tag == "a" and self.href:
            self.links.append((urljoin(SOURCE_PAGE, self.href), self.label.lower()))
            self.href = None


def download(url):
    class NoRedirect(HTTPRedirectHandler):
        def redirect_request(self, req, fp, code, msg, headers, newurl):
            return None

    opener = build_opener(NoRedirect())
    # Check every redirect before following it; only read official public pages.
    for _ in range(5):
        parsed = urlparse(url)
        if parsed.scheme != "https" or parsed.hostname not in {"www.vic.gov.au", "vic.gov.au"}:
            raise ValueError("Calendar download moved outside the approved government website.")
        try:
            response = opener.open(Request(url, headers={"User-Agent": "VictorianRosterCalendar/1.0"}), timeout=15)
        except HTTPError as exc:
            if exc.code in (301, 302, 303, 307, 308):
                url = urljoin(url, exc.headers["Location"])
                continue
            raise
        with response:
            content = response.read(2_000_001)
            if len(content) > 2_000_000:
                raise ValueError("Government calendar response is unexpectedly large.")
            return content.decode("utf-8-sig")
    raise ValueError("Too many calendar download redirects.")


def discover(html):
    parser = CalendarLinks()
    parser.feed(html)
    found = {}
    for kind, phrase in (("public", "public holiday"), ("school", "school term")):
        matches = {url for url, label in parser.links
                   if urlparse(url).path.lower().endswith(".ics")
                   and phrase in (label + " " + url.lower().replace("-", " "))}
        if len(matches) != 1:
            raise ValueError(f"Could not identify a unique official {kind} calendar. Manual review needed.")
        found[kind] = matches.pop()
    return found


def parse_events(text):
    if "BEGIN:VCALENDAR" not in text or "END:VCALENDAR" not in text:
        raise ValueError("The source did not contain a complete calendar.")
    unfolded = re.sub(r"\r?\n[ \t]", "", text)
    events = []
    for block in re.findall(r"BEGIN:VEVENT\s*(.*?)END:VEVENT", unfolded, re.S):
        props = {}
        depth = 0
        for line in block.splitlines():
            if line.startswith("BEGIN:"):
                depth += 1
            elif line.startswith("END:"):
                depth -= 1
            elif depth == 0 and ":" in line:
                key, value = line.split(":", 1)
                key = key.split(";", 1)[0]
                if key in props:
                    raise ValueError("Duplicate calendar property; manual review needed.")
                props[key] = value.strip()
        if any(key in props for key in ("RRULE", "RDATE", "EXDATE", "RECURRENCE-ID")):
            raise ValueError("Recurring events are not supported by this holiday reader.")
        if not re.fullmatch(r"\d{8}", props.get("DTSTART", "")):
            raise ValueError("Expected an all-day government calendar event.")
        start = datetime.strptime(props["DTSTART"], "%Y%m%d").date()
        end = datetime.strptime(props.get("DTEND", ""), "%Y%m%d").date()
        if end != start + timedelta(days=1) or not props.get("SUMMARY"):
            raise ValueError("Unexpected government calendar event format.")
        props["date"] = start
        events.append(props)
    if not events or text.count("BEGIN:VEVENT") != len(events):
        raise ValueError("Calendar was empty or incomplete.")
    return events


def public_records(text, years):
    result, warnings = [], []
    for event in parse_events(text):
        day = event["date"]
        if day.year not in years:
            continue
        name = event["SUMMARY"].replace("\\,", ",").replace("\\n", " ")
        notes = event.get("DESCRIPTION", "")
        provisional = bool(re.search(r"subject to|typically|exact date|TBA|to be confirmed|tentative",
                                     name + " " + notes + " " + event.get("STATUS", ""), re.I))
        if "Grand Final" in name and day.weekday() != 4:
            provisional = True
        if event.get("STATUS") == "CANCELLED":
            raise ValueError("The source includes a cancelled holiday. Manual review needed.")
        if provisional:
            warnings.append(f"{day.year}: {name} is provisional in the source; existing entry will be retained.")
        result.append({"name": name, "start": day.isoformat(), "end": day.isoformat(),
                       "provisional": provisional, "notes": notes})
    for year in years:
        rows = [r for r in result if r["start"].startswith(str(year))]
        if not 11 <= len(rows) <= 20 or len({r["start"] for r in rows}) != len(rows):
            raise ValueError(f"Public holiday coverage for {year} is missing, duplicated or incomplete.")
        for anchor in ("New Year", "Australia Day", "Good Friday", "Christmas", "Boxing"):
            if not any(anchor.lower() in r["name"].lower() for r in rows):
                raise ValueError(f"Missing expected holiday in {year}: {anchor}.")
    return sorted(result, key=lambda r: r["start"]), warnings


def school_records(text, years):
    terms = {}
    for event in parse_events(text):
        match = re.fullmatch(r"Term ([1-4]) (starts|ends)( \(government schools\))?", event["SUMMARY"])
        if not match:
            raise ValueError("Unrecognised school calendar event. Manual review needed.")
        term, boundary, students = match.groups()
        key = (event["date"].year, int(term), boundary)
        # Term 1 has a teacher start and a separate government-school student start.
        priority = 1 if students else 0
        if key in terms and priority == terms[key][0]:
            raise ValueError("Duplicate school term boundary.")
        if key not in terms or priority > terms[key][0]:
            terms[key] = (priority, event["date"])
    result = []
    for year in years:
        for term, season in enumerate(("Autumn", "Winter", "Spring", "Summer"), 1):
            next_key = (year, term + 1, "starts") if term < 4 else (year + 1, 1, "starts")
            try:
                start = terms[(year, term, "ends")][1] + timedelta(days=1)
                end = terms[next_key][1] - timedelta(days=1)
            except KeyError as exc:
                raise ValueError(f"School term coverage for {year} is incomplete.") from exc
            if not 7 <= (end - start).days <= 65:
                raise ValueError("Unexpected school holiday duration; manual review needed.")
            label = str(year) if term < 4 else f"{year}-{year + 1}"
            result.append({"name": f"{label} {season} School Holidays", "start": start.isoformat(),
                           "end": end.isoformat(), "provisional": False})
    return result


def fetch_proposal(today, fetch=download):
    urls = discover(fetch(SOURCE_PAGE))
    years = [today.year, today.year + 1]
    public, warnings = public_records(fetch(urls["public"]), years)
    school = school_records(fetch(urls["school"]), years)
    return {"public": public, "school": school, "years": years, "sources": urls,
            "checked_at": stamp(), "warnings": warnings}


def baseline(public, school):
    return {"public": [{"name": name, "start": day.isoformat(), "end": day.isoformat()}
                       for day, name in sorted(public.items())],
            "school": [{"name": r["name"], "start": r["start"].isoformat(), "end": r["end"].isoformat()}
                       for r in school]}


def candidate_calendar(active, proposal):
    result = copy.deepcopy(active)
    years = proposal["years"]
    for kind in ("public", "school"):
        retained = [r for r in active[kind] if date.fromisoformat(r["start"]).year not in years]
        if kind == "public":
            # The official feed may publish placeholders for AFL dates. Never apply them.
            def identity(row):
                name = re.sub(r"\([^)]*\)", "", row["name"].lower())
                name = re.sub(r"\bthe\b", "", name)
                return (date.fromisoformat(row["start"]).year, " ".join(name.split()))
            blocked = {identity(r) for r in proposal[kind] if r["provisional"]}
            retained += [r for r in active[kind] if identity(r) in blocked]
        result[kind] = sorted(retained + [r for r in proposal[kind] if not r["provisional"]],
                              key=lambda r: r["start"])
    return result


def changes(active, candidate):
    rows = []
    for kind in ("public", "school"):
        before = {(r["start"], r["end"]): r["name"] for r in active[kind]}
        after = {(r["start"], r["end"]): r["name"] for r in candidate[kind]}
        for key in sorted(before.keys() | after.keys()):
            if before.get(key) != after.get(key):
                rows.append({"Type": kind, "From": key[0], "To": key[1],
                             "Saved": before.get(key, "—"), "Published": after.get(key, "—")})
    return rows


def fingerprint(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True).encode()).hexdigest()


def check_due(state, today):
    previous = state.get("last_attempt")
    if not previous:
        return True
    last = date.fromisoformat(previous[:10])
    interval = 1 if state.get("error") else (7 if today.month in (10, 11, 12, 1) else 30)
    return (today - last).days >= interval


def check_state(state, active, today, fetch=download):
    updated = copy.deepcopy(state)
    updated["last_attempt"] = today.isoformat()
    try:
        proposal = fetch_proposal(today, fetch)
        proposal["base_hash"] = fingerprint(active)
        updated.update(pending=proposal, last_success=proposal["checked_at"], error=None)
    except Exception as exc:
        # A failed check must not leave a stale proposal available for approval.
        updated.update(pending=None, error=f"Holiday dates could not be verified. {type(exc).__name__}. Manual review needed.")
    return updated


def approve(state, active, actor):
    proposal = state.get("pending")
    if not proposal or proposal["base_hash"] != fingerprint(active):
        raise ValueError("Calendar changed since this preview. Check the sources again.")
    updated = copy.deepcopy(state)
    updated["previous_calendar"] = active
    updated["calendar"] = candidate_calendar(active, proposal)
    updated["approved_at"] = stamp()
    updated["approved_by"] = actor
    updated["sources"] = proposal["sources"]
    updated["warnings"] = proposal["warnings"]
    updated["pending"] = None
    return updated
