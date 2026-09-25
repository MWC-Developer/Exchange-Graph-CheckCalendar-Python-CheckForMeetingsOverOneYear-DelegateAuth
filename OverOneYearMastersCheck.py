# OverOneYearMastersCheck.py

"""Find calendar master events that span over one year or have no end date.

Required Microsoft Graph application permission: Calendars.Read (admin consent).
Set GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, and GRAPH_USER_ID in the
environment before running this script. No third-party Python packages are required.

This script intentionally uses /users/{id}/events and never uses calendarView. The
/events endpoint returns master records without expanding recurring occurrences.
"""

from __future__ import annotations

import csv
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlencode
from urllib.request import Request, urlopen


TENANT_ID = os.environ.get("GRAPH_TENANT_ID", "")
CLIENT_ID = os.environ.get("GRAPH_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("GRAPH_CLIENT_SECRET", "")
USER_ID = os.environ.get("GRAPH_USER_ID", "")

GRAPH_BASE_URI = "https://graph.microsoft.com/v1.0"
PAGE_SIZE = 500
OUTPUT_PATH = Path(__file__).resolve().parent / "OverOneYearMastersCheck.CSV"
CSV_FIELDS = [
    "MasterEventId",
    "ICalUid",
    "Organizer",
    "Subject",
    "StartTime",
    "EndTime",
    "CreatedDateTime",
]

JsonObject = dict[str, Any]


def log(message: str) -> None:
    print(message, flush=True)


def validate_configuration() -> None:
    settings = {
        "GRAPH_TENANT_ID": TENANT_ID,
        "GRAPH_CLIENT_ID": CLIENT_ID,
        "GRAPH_CLIENT_SECRET": CLIENT_SECRET,
        "GRAPH_USER_ID": USER_ID,
    }
    missing = [name for name, value in settings.items() if not value.strip()]
    if missing:
        raise ValueError(f"Set the following environment variables: {', '.join(missing)}")
    if not 1 <= PAGE_SIZE <= 999:
        raise ValueError("PAGE_SIZE must be between 1 and 999.")


def graph_request(
    uri: str,
    *,
    method: str,
    headers: dict[str, str] | None = None,
    body: bytes | None = None,
    maximum_attempts: int = 5,
) -> JsonObject:
    transient_statuses = {429, 500, 502, 503, 504}

    for attempt in range(1, maximum_attempts + 1):
        request = Request(uri, data=body, headers=headers or {}, method=method)
        try:
            with urlopen(request, timeout=120) as response:
                result = json.loads(response.read().decode("utf-8"))
                if not isinstance(result, dict):
                    raise RuntimeError(f"Expected a JSON object from {uri}")
                return result
        except HTTPError as error:
            if error.code not in transient_statuses or attempt == maximum_attempts:
                detail = error.read().decode("utf-8", errors="replace")
                raise RuntimeError(
                    f"Graph request failed with HTTP {error.code}: {detail}"
                ) from error
            retry_after = error.headers.get("Retry-After")
            delay = int(retry_after) if retry_after and retry_after.isdigit() else min(2**attempt, 30)
            log(
                f"Graph returned HTTP {error.code}; retrying in {delay} seconds "
                f"(attempt {attempt} of {maximum_attempts})."
            )
            time.sleep(delay)
        except URLError as error:
            if attempt == maximum_attempts:
                raise RuntimeError(f"Graph request failed: {error.reason}") from error
            delay = min(2**attempt, 30)
            log(
                f"Graph request failed: {error.reason}; retrying in {delay} seconds "
                f"(attempt {attempt} of {maximum_attempts})."
            )
            time.sleep(delay)

    raise RuntimeError(f"Graph request failed after {maximum_attempts} attempts: {uri}")


def get_access_token() -> str:
    token_uri = (
        f"https://login.microsoftonline.com/{quote(TENANT_ID, safe='')}"
        "/oauth2/v2.0/token"
    )
    body = urlencode(
        {
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        }
    ).encode("ascii")
    response = graph_request(
        token_uri,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        body=body,
    )
    access_token = response.get("access_token")
    if not isinstance(access_token, str) or not access_token:
        raise RuntimeError("The OAuth response did not contain an access token.")
    return access_token


def iter_master_events(access_token: str) -> Iterator[JsonObject]:
    selected_fields = (
        "id,iCalUId,organizer,subject,start,end,createdDateTime,type,recurrence"
    )
    query = urlencode({"$select": selected_fields, "$top": PAGE_SIZE})
    encoded_user_id = quote(USER_ID, safe="")
    next_uri: str | None = f"{GRAPH_BASE_URI}/users/{encoded_user_id}/events?{query}"
    page_number = 0

    while next_uri:
        page_number += 1
        log(f"Reading master event page {page_number}...")
        response = graph_request(
            next_uri,
            method="GET",
            headers={
                "Authorization": f"Bearer {access_token}",
                "Accept": "application/json",
                "Prefer": f"odata.maxpagesize={PAGE_SIZE}",
            },
        )
        items = response.get("value")
        if not isinstance(items, list):
            raise RuntimeError("Graph response 'value' property was not an array.")
        for event in items:
            if isinstance(event, dict) and event.get("type") in {
                "seriesMaster",
                "singleInstance",
            }:
                yield event
        next_link = response.get("@odata.nextLink")
        next_uri = next_link if isinstance(next_link, str) and next_link else None


def parse_graph_datetime(value: Any) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None


def add_one_year(value: datetime) -> datetime:
    try:
        return value.replace(year=value.year + 1)
    except ValueError:
        return value.replace(month=2, day=28, year=value.year + 1)


def recurrence_range(event: JsonObject) -> JsonObject | None:
    recurrence = event.get("recurrence")
    if not isinstance(recurrence, dict):
        return None
    range_value = recurrence.get("range")
    return range_value if isinstance(range_value, dict) else None


def recurrence_has_no_end(event: JsonObject) -> bool:
    range_value = recurrence_range(event)
    if range_value is None:
        return False
    if str(range_value.get("type") or "").casefold() == "noend":
        return True
    end_date = parse_graph_datetime(range_value.get("endDate"))
    return end_date is None or end_date.date() == datetime.min.date()


def event_boundaries(event: JsonObject) -> tuple[datetime | None, datetime | None]:
    range_value = recurrence_range(event)
    if range_value is not None:
        return (
            parse_graph_datetime(range_value.get("startDate")),
            parse_graph_datetime(range_value.get("endDate")),
        )

    start = event.get("start")
    end = event.get("end")
    start_value = start.get("dateTime") if isinstance(start, dict) else None
    end_value = end.get("dateTime") if isinstance(end, dict) else None
    return parse_graph_datetime(start_value), parse_graph_datetime(end_value)


def event_matches(event: JsonObject) -> bool:
    if recurrence_has_no_end(event):
        return True
    start, end = event_boundaries(event)
    return bool(start and end and end > add_one_year(start))


def graph_datetime_text(value: Any) -> str:
    if not isinstance(value, dict) or not value.get("dateTime"):
        return ""
    date_time = str(value["dateTime"])
    time_zone = str(value.get("timeZone") or "")
    return f"{date_time} [{time_zone}]" if time_zone else date_time


def boundary_text(event: JsonObject, boundary: str) -> str:
    range_value = recurrence_range(event)
    if range_value is not None:
        if boundary == "End" and recurrence_has_no_end(event):
            return ""
        date_key = "startDate" if boundary == "Start" else "endDate"
        date_text = str(range_value.get(date_key) or "")
        time_value = event.get("start" if boundary == "Start" else "end")
        if isinstance(time_value, dict) and time_value.get("dateTime"):
            time_text = str(time_value["dateTime"]).split("T", 1)
            if len(time_text) == 2:
                zone = str(time_value.get("timeZone") or "")
                result = f"{date_text}T{time_text[1]}" if date_text else ""
                return f"{result} [{zone}]" if result and zone else result
        return date_text
    return graph_datetime_text(event.get("start" if boundary == "Start" else "end"))


def organizer_text(event: JsonObject) -> str:
    organizer = event.get("organizer")
    if not isinstance(organizer, dict):
        return ""
    email_address = organizer.get("emailAddress")
    if not isinstance(email_address, dict):
        return ""
    name = str(email_address.get("name") or "")
    address = str(email_address.get("address") or "")
    if name and address:
        return f"{name} <{address}>"
    return address or name


def csv_safe(value: Any) -> str:
    text = str(value or "")
    if text.startswith(("=", "+", "-", "@")):
        return f"'{text}"
    return text


def event_row(event: JsonObject) -> dict[str, str]:
    return {
        "MasterEventId": csv_safe(event.get("id")),
        "ICalUid": csv_safe(event.get("iCalUId")),
        "Organizer": csv_safe(organizer_text(event)),
        "Subject": csv_safe(event.get("subject")),
        "StartTime": csv_safe(boundary_text(event, "Start")),
        "EndTime": csv_safe(boundary_text(event, "End")),
        "CreatedDateTime": csv_safe(event.get("createdDateTime")),
    }


def write_csv(rows: list[dict[str, str]]) -> None:
    with OUTPUT_PATH.open("w", encoding="utf-8-sig", newline="") as output_file:
        writer = csv.DictWriter(
            output_file,
            fieldnames=CSV_FIELDS,
            quoting=csv.QUOTE_ALL,
        )
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    try:
        validate_configuration()
        log("Acquiring Microsoft Graph application token...")
        access_token = get_access_token()
        rows = [event_row(event) for event in iter_master_events(access_token) if event_matches(event)]
        rows.sort(key=lambda row: (row["Subject"].casefold(), row["MasterEventId"]))
        write_csv(rows)
        log(f"Matched master events: {len(rows)}")
        log(f"CSV written to: {OUTPUT_PATH}")
        return 0
    except (OSError, RuntimeError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
