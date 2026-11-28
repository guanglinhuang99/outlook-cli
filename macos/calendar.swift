import EventKit
import Foundation

private struct Request: Decodable {
    let action: String
    let fromDate: String?
    let endDateExclusive: String?
    let timeZone: String?
    let calendarIDs: [String]?
    let includeDetails: Bool?
    let includeCanceled: Bool?
    let query: String?

    enum CodingKeys: String, CodingKey {
        case action, query
        case fromDate = "from_date"
        case endDateExclusive = "end_date_exclusive"
        case timeZone = "time_zone"
        case calendarIDs = "calendar_ids"
        case includeDetails = "include_details"
        case includeCanceled = "include_canceled"
    }
}

private func authorization() -> String {
    // Full access and the legacy authorized spelling both use value 3.
    switch EKEventStore.authorizationStatus(for: .event).rawValue {
    case 0: return "not_determined"
    case 1: return "restricted"
    case 2: return "denied"
    case 3: return "full_access"
    case 4: return "write_only"
    default: return "unknown"
    }
}

private func failure(_ code: String, _ message: String) -> [String: Any] {
    ["error": ["code": code, "message": message]]
}

private func dayFormatter(in timeZone: TimeZone) -> DateFormatter {
    let formatter = DateFormatter()
    formatter.locale = Locale(identifier: "en_US_POSIX")
    formatter.calendar = Calendar(identifier: .gregorian)
    formatter.timeZone = timeZone
    formatter.dateFormat = "yyyy-MM-dd"
    formatter.isLenient = false
    return formatter
}

private func dayStart(_ value: String, timeZone: TimeZone, formatter: DateFormatter) -> Date? {
    let fields = value.split(separator: "-").compactMap { Int($0) }
    guard fields.count == 3 else { return nil }
    var calendar = Calendar(identifier: .gregorian)
    calendar.timeZone = timeZone
    guard let day = calendar.date(from: DateComponents(year: fields[0], month: fields[1], day: fields[2])),
          formatter.string(from: day) == value else { return nil }
    return calendar.startOfDay(for: day)
}

private func calendarRow(_ calendar: EKCalendar) -> [String: Any] {
    [
        "id": calendar.calendarIdentifier,
        "title": String(calendar.title.prefix(1024)),
        "source_id": calendar.source.sourceIdentifier,
        "source_title": String(calendar.source.title.prefix(1024))
    ]
}

private func eventRow(
    _ event: EKEvent, formatter: ISO8601DateFormatter, includeDetails: Bool
) -> [String: Any]? {
    let identifier = event.calendarItemIdentifier
    guard let calendar = event.calendar, let start = event.startDate,
          let end = event.endDate, !identifier.isEmpty else { return nil }
    let timeZone = event.timeZone ?? TimeZone.current
    let location = event.location ?? ""
    var row: [String: Any] = [
        "id": identifier,
        "calendar_id": calendar.calendarIdentifier,
        "title": String((event.title ?? "").prefix(1024)),
        "start_at": formatter.string(from: start),
        "end_at": formatter.string(from: end),
        "all_day": event.isAllDay,
        "time_zone": timeZone.identifier,
        "recurring": event.hasRecurrenceRules || event.isDetached,
        "canceled": event.status == .canceled
    ]
    if let occurrence = event.occurrenceDate {
        row["occurrence_at"] = formatter.string(from: occurrence)
    }
    if event.isAllDay {
        let days = dayFormatter(in: timeZone)
        row["start_date"] = days.string(from: start)
        row["end_date_exclusive"] = days.string(from: end)
    }
    if includeDetails {
        let notes = event.notes ?? ""
        let attendees = event.attendees ?? []
        row["location"] = String(location.prefix(1024))
        row["notes"] = String(notes.prefix(4096))
        row["url"] = String((event.url?.absoluteString ?? "").prefix(2048))
        row["attendees"] = attendees.prefix(50).map { attendee in
            [
                "name": String((attendee.name ?? "").prefix(256)),
                "url": String(attendee.url.absoluteString.prefix(1024))
            ]
        }
        row["details_truncated"] = notes.count > 4096 || location.count > 1024
            || (event.url?.absoluteString.count ?? 0) > 2048 || attendees.count > 50
            || attendees.prefix(50).contains { attendee in
                (attendee.name?.count ?? 0) > 256 || attendee.url.absoluteString.count > 1024
            }
    }
    return row
}

private func agenda(_ request: Request, store: EKEventStore) -> [String: Any] {
    guard let zone = request.timeZone.flatMap(TimeZone.init(identifier:)),
          let fromDate = request.fromDate, let endDate = request.endDateExclusive else {
        return failure("INVALID_ARGUMENT", "Agenda requires dates and a supported time zone")
    }
    let days = dayFormatter(in: zone)
    guard let start = dayStart(fromDate, timeZone: zone, formatter: days),
          let end = dayStart(endDate, timeZone: zone, formatter: days), start < end else {
        return failure("INVALID_ARGUMENT", "Agenda dates do not exist in the requested time zone")
    }
    let localCalendars = store.calendars(for: .event)
    let selectedCalendars: [EKCalendar]
    if let ids = request.calendarIDs, !ids.isEmpty {
        let selected = Set(ids)
        selectedCalendars = localCalendars.filter { selected.contains($0.calendarIdentifier) }
        guard selectedCalendars.count == selected.count else {
            return failure("CALENDAR_NOT_FOUND", "One or more calendar IDs do not exist in this Mac's EventKit store")
        }
    } else {
        selectedCalendars = localCalendars
    }
    guard !selectedCalendars.isEmpty else { return ["items": []] }

    let formatter = ISO8601DateFormatter()
    formatter.formatOptions = [.withInternetDateTime, .withFractionalSeconds]
    let predicate = store.predicateForEvents(withStart: start, end: end, calendars: selectedCalendars)
    var rows: [[String: Any]] = []
    var examined = 0
    var invalid = false
    store.enumerateEvents(matching: predicate) { event, stop in
        examined += 1
        if examined > 4_096 {
            stop.pointee = true
            return
        }
        guard let eventStart = event.startDate, let eventEnd = event.endDate else {
            invalid = true
            return
        }
        guard eventStart < end && (eventEnd > start || eventStart >= start) else { return }
        if request.includeCanceled != true && event.status == .canceled { return }
        if let query = request.query {
            let found = [event.title ?? "", event.location ?? ""].contains {
                $0.range(of: query, options: [.caseInsensitive, .diacriticInsensitive]) != nil
            }
            if !found { return }
        }
        if let row = eventRow(event, formatter: formatter, includeDetails: request.includeDetails == true) {
            rows.append(row)
        } else {
            invalid = true
        }
    }
    if examined > 4_096 {
        return failure("CALENDAR_QUERY_TOO_LARGE", "Agenda exceeds 4096 EventKit occurrences. Narrow the date range or select calendar IDs")
    }
    if invalid {
        return failure("CALENDAR_INVALID_EVENT", "EventKit returned an occurrence without a date or identifier")
    }
    return ["items": rows]
}

@main
private struct CalendarReader {
    static func main() async {
        let output: [String: Any]
        do {
            let request = try JSONDecoder().decode(Request.self, from: FileHandle.standardInput.readDataToEndOfFile())
            output = await run(request)
        } catch {
            output = failure("INVALID_ARGUMENT", "Invalid EventKit request")
        }
        if let json = try? JSONSerialization.data(withJSONObject: output, options: [.sortedKeys]) {
            FileHandle.standardOutput.write(json)
        }
    }

    static func run(_ request: Request) async -> [String: Any] {
        let current = authorization()
        if request.action == "status" {
            return ["authorization": current]
        }
        if request.action == "authorize" {
            if current == "full_access" { return ["authorization": current] }
            if current == "denied" || current == "restricted" {
                return failure("CALENDAR_PERMISSION_DENIED", "Allow Calendar access for the Outlook CLI host in macOS Privacy & Security")
            }
            guard #available(macOS 14, *) else {
                return failure("CALENDAR_UNAVAILABLE", "Calendar authorization requires macOS 14 or later")
            }
            let store = EKEventStore()
            do {
                _ = try await store.requestFullAccessToEvents()
            } catch {
                // The resulting TCC state distinguishes denial from an unavailable prompt.
            }
            switch authorization() {
            case "full_access":
                return ["authorization": "full_access"]
            case "denied", "restricted":
                return failure("CALENDAR_PERMISSION_DENIED", "Full Calendar access was not granted")
            default:
                return failure("CALENDAR_AUTH_UNAVAILABLE", "EventKit did not grant full Calendar access in the current macOS host session")
            }
        }
        if current == "denied" || current == "restricted" {
            return failure("CALENDAR_PERMISSION_DENIED", "Allow Calendar access for the Outlook CLI host in macOS Privacy & Security")
        }
        guard current == "full_access" else {
            return failure("CALENDAR_AUTH_REQUIRED", "Full EventKit Calendar access is required. Run outlook-local calendar authorize from a logged-in macOS session")
        }
        let store = EKEventStore()
        switch request.action {
        case "calendars":
            let calendars = store.calendars(for: .event)
            guard calendars.count <= 4_096 else {
                return failure("CALENDAR_QUERY_TOO_LARGE", "EventKit returned more than 4096 calendars")
            }
            return ["items": calendars.map(calendarRow)]
        case "agenda":
            return agenda(request, store: store)
        default:
            return failure("INVALID_ARGUMENT", "Unknown EventKit action")
        }
    }
}
