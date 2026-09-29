"""Manager controls and durable storage for reviewed holiday updates."""
from pathlib import Path
from datetime import date, timedelta
import json
import os
import tempfile

from holiday_sources import (SOURCE_PAGE, approve, baseline, candidate_calendar,
                             changes, check_due, check_state, fingerprint)


class CalendarStore:
    def __init__(self, directory, db=None):
        self.path = Path(directory) / "holiday_calendar.json"
        self.db = db

    def load(self):
        if self.db is not None:
            doc = self.db.collection("system").document("holiday_calendar").get()
            return doc.to_dict() if doc.exists else {}
        if not self.path.exists():
            return {}
        return json.loads(self.path.read_text(encoding="utf-8"))

    def save(self, previous, updated):
        if self.db is not None:
            from firebase_admin import firestore
            ref = self.db.collection("system").document("holiday_calendar")

            @firestore.transactional
            def commit(transaction):
                doc = ref.get(transaction=transaction)
                current = doc.to_dict() if doc.exists else {}
                if fingerprint(current) != fingerprint(previous):
                    raise ValueError("Another operator updated this calendar. Reload before continuing.")
                transaction.set(ref, updated)
            commit(self.db.transaction())
            return
        self.path.parent.mkdir(parents=True, exist_ok=True)
        lock = self.path.with_suffix(".lock")
        # Exclusive lock prevents two local sessions overwriting each other's approval.
        fd = os.open(str(lock), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        temp = None
        try:
            if fingerprint(self.load()) != fingerprint(previous):
                raise ValueError("Another operator updated this calendar. Reload before continuing.")
            with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=self.path.parent,
                                             delete=False) as handle:
                temp = handle.name
                json.dump(updated, handle, indent=2)
            os.replace(temp, self.path)
            temp = None
        finally:
            os.close(fd)
            lock.unlink()
            if temp:
                Path(temp).unlink(missing_ok=True)


def effective_calendar(state, default_calendar):
    active = state.get("calendar", default_calendar)
    # Fail clearly on corrupted storage rather than silently resetting approved dates.
    for kind in ("public", "school"):
        for row in active[kind]:
            if not row["name"] or date.fromisoformat(row["end"]) < date.fromisoformat(row["start"]):
                raise ValueError("Saved holiday calendar is invalid.")
    return active


def as_roster_dates(active):
    return ({date.fromisoformat(r["start"]): r["name"] for r in active["public"]},
            [{"name": r["name"], "start": date.fromisoformat(r["start"]),
              "end": date.fromisoformat(r["end"])} for r in active["school"]])


def render_panel(st, store, defaults, today, actor, finalized_rosters):
    with st.expander("Holiday calendar · official updates", expanded=False):
        st.caption("Victoria public holidays and government-school holidays. Local substitute holidays "
                   "and individual school pupil-free days need a separate check.")
        st.markdown(f"[View official calendar downloads]({SOURCE_PAGE})")
        try:
            state = store.load()
            active = effective_calendar(state, defaults)
        except Exception:
            st.error("Saved holiday calendar could not be loaded. Updates are disabled; check storage connectivity.")
            return
        automatic = st.checkbox("Check automatically when a manager opens the dashboard",
                                value=state.get("automatic", True), key="holiday_auto_check")
        st.caption("Monthly, or weekly October–January. Closed-app background checking is not configured.")
        try:
            if automatic != state.get("automatic", True):
                updated = dict(state, automatic=automatic)
                store.save(state, updated)
                state = updated
            requested = st.button("Check official dates now", key="holiday_check_now")
            if requested or (automatic and check_due(state, today)):
                with st.spinner("Checking the official Victorian calendars…"):
                    updated = check_state(state, active, today)
                    store.save(state, updated)
                    state = updated
        except Exception:
            st.error("The check could not be saved. Reload and try again; your active calendar was not changed.")
            return
        st.caption(f"Last successful check: {state.get('last_success', 'Never')} · "
                   f"Last attempt: {state.get('last_attempt', 'Never')}")
        if state.get("error"):
            st.warning(state["error"])
        if not state.get("calendar"):
            st.info("Using the bundled calendar. Official date updates take effect only after review below.")
        if today.month >= 10:
            next_year = str(today.year + 1)
            if any(not any(r["start"].startswith(next_year) for r in active[kind])
                   for kind in ("public", "school")):
                st.warning("Next year's saved holiday calendar is incomplete. Check and review official dates.")
        for warning in state.get("warnings", []):
            st.warning(warning)
        upcoming = [{"Type": kind, "Event": r["name"], "Start": r["start"], "End": r["end"]}
                    for kind in ("public", "school") for r in active[kind]
                    if date.fromisoformat(r["end"]) >= today
                    and date.fromisoformat(r["start"]) <= today + timedelta(days=90)]
        if upcoming:
            st.write("Upcoming 90 days · saved calendar")
            st.dataframe(sorted(upcoming, key=lambda r: r["Start"]), hide_index=True, use_container_width=True)
        pending = state.get("pending")
        if not pending:
            return
        candidate = candidate_calendar(active, pending)
        delta = changes(active, candidate)
        for warning in pending["warnings"]:
            st.warning(warning)
        st.caption("School breaks use the published student start dates; teacher start dates may be earlier.")
        if delta:
            st.write("Proposed changes · not applied")
            st.dataframe(delta, hide_index=True, use_container_width=True)
            affected = []
            for roster in finalized_rosters:
                raw = str(roster.get("date_str", ""))
                try:
                    start = date.fromisoformat(raw)
                except ValueError:
                    continue
                if any(date.fromisoformat(row["From"]) <= start + timedelta(days=6)
                       and date.fromisoformat(row["To"]) >= start for row in delta):
                    affected.append(roster.get("label", raw))
            if affected:
                st.warning("Published roster weeks requiring review: " + ", ".join(affected))
            st.info("Applying updates changes holiday detection and live wage calculations. "
                    "Existing roster shifts and exported files are not rewritten.")
        else:
            st.success("No date or name changes found in the checked years.")
        acknowledged = st.checkbox("I have reviewed these dates, source warnings and local holiday applicability.",
                                   key="holiday_review_" + fingerprint(pending)[:16])
        if st.button("Apply reviewed calendar" if delta else "Confirm checked calendar",
                     disabled=not acknowledged, key="holiday_apply"):
            try:
                updated = approve(state, active, actor)
                store.save(state, updated)
            except Exception:
                st.error("Calendar was not applied. Another update or storage failure may have occurred. Reload and check again.")
            else:
                st.success("Reviewed holiday calendar saved.")
                st.rerun()
