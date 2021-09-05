import Flutter
import UIKit
import EventKit
import Foundation

extension Date {
    var millisecondsSinceEpoch: Double { return self.timeIntervalSince1970 * 1000.0 }
}

extension EKParticipant {
    var emailAddress: String? {
        return self.value(forKey: "emailAddress") as? String
    }
}

/// Defines frequencies for recurrence rules.
///
/// - daily: Indicates a daily recurrence rule.
/// - weekly: Indicates a weekly recurrence rule.
/// - monthly: Indicates a monthly recurrence rule.
/// - yearly: Indicates a yearly recurrence rule.
public enum RWMRecurrenceFrequency: Int, Codable {
    case daily = 0
    case weekly = 1
    case monthly = 2
    case yearly = 3
}

public enum RWMWeekday: Int, Codable {
    case sunday = 1
    case monday = 2
    case tuesday = 3
    case wednesday = 4
    case thursday = 5
    case friday = 6
    case saturday = 7
}

public class RWMRuleParser {
    private lazy var untilFormat: DateFormatter = {
        let df = DateFormatter()
        df.locale = Locale(identifier: "en_US_POSIX")
        df.timeZone = TimeZone(secondsFromGMT: 0)
        df.dateFormat = "yyyyMMdd'T'HHmmssX"
        return df
    }()

    public init() {
    }

    /// Compares two RRULE strings to see if they have the same components. The components do not need to be in the
    /// same order. Any `UNTIL` clause is ignored since the date can be in a different format.
    ///
    /// - Parameters:
    ///   - left: The first RRULE string.
    ///   - right: The second RRULE string.
    /// - Returns: `true` if the two rules have the same components, ignoring order and any `UNTIL` clause. `false` if different.
    public func compare(rule left: String, to right: String) -> Bool {
        var leftParts = split(rule: left).sorted()
        var rightParts = split(rule: right).sorted()
        if leftParts.first(where: { $0.hasPrefix("UNTIL") }) != nil && rightParts.first(where: { $0.hasPrefix("UNTIL")}) != nil {
            leftParts = leftParts.filter { !$0.hasPrefix("UNTIL") }
            rightParts = leftParts.filter { !$0.hasPrefix("UNTIL") }
        }

        return leftParts == rightParts
    }

    private func split(rule: String) -> [String] {
        var r = rule.uppercased()
        if r.hasPrefix("RRULE:") {
            r.removeFirst(6)
        }

        let parts = r.components(separatedBy: ";")

        return parts
    }

    /// Parses an RRULE string returning a `RWMRecurrenceRule`.
    ///
    /// Valid strings:
    ///   - The RRULE string may optionally begin with `RRULE:`.
    ///   - There must be 1 `FREQ=` followed by either `DAILY`, `WEEKLY`, `MONTHLY`, `YEARLY`.
    ///   - There may be 1 `COUNT=` followed by a positive integer.
    ///   - There may be 1 `UNTIL=` followed by a date. The date may be in one of these formats: "yyyyMMdd'T'HHmmssX", "yyyyMMdd'T'HHmmss", "'TZID'=VV:yyyyMMdd'T'HHmmss", "yyyyMMdd".
    ///   - Only 1 of either `COUNT` or `UNTIL` is allowed, not both.
    ///   - There may be 1 `INTERVAL=` followed by a positive integer.
    ///   - There may be 1 `BYMONTH=` followed by a comma separated list of 1 or more month numbers in the range 1 to 12, or -12 to -1.
    ///   - There may be 1 `BYDAY=` followed by a comma separated list of 1 or more days of the week, each optionally preceded by a week number. Days of the week must be `SU`, `MO`, `TU`, `WE`, `TH`, `FR`, or `SA`. Week numbers must be in the range 1 to 5 or -5 to -1.
    ///   - There may be 1 `BYMONTHDAY=` followed by a comma separated list of days of the month in the range 1 to 31 or -31 to -1.
    ///   - There may be 1 `BYYEARDAY=` followed by a comma separated list of days of the year in the range 1 to 366 or -366 to -1.
    ///   - There may be 1 `BYWEEKNO=` followed by a comma separated list of week numbers in the range 1 to 53 or -53 to -1.
    ///   - There may be 1 `WKST=` followed by a day of the week. Days of the week must be `SU`, `MO`, `TU`, `WE`, `TH`, `FR`, or `SA`.
    ///   - There may be 1 `BYSETPOS=` following by a comma separated list of positive integers.
    ///   - Each clause must be separated by a semicolon (`;`). No trailing semicolon should be used.
    ///
    /// - Parameter rule: The RRULE string.
    /// - Returns: The resulting recurrence rule. If the RRULE string is invalid in any way, the result is `nil`.
    public func parse(rule: String) -> RWMRecurrenceRule? {
        var frequency: RWMRecurrenceFrequency? = nil
        var interval: Int? = nil
        var firstDayOfTheWeek: RWMWeekday? = nil
        var daysOfTheWeek: [RWMRecurrenceDayOfWeek]? = nil
        var daysOfTheMonth: [Int]? = nil
        var daysOfTheYear: [Int]? = nil
        var weeksOfTheYear: [Int]? = nil
        var monthsOfTheYear: [Int]? = nil
        var setPositions: [Int]? = nil
        var recurrenceEnd: RWMRecurrenceEnd? = nil

        let parts = split(rule: rule)
        for part in parts {
            let varval = part.components(separatedBy: "=")
            guard varval.count == 2 else { return nil }

            switch varval[0] {
            case "FREQ":
                guard frequency == nil else { return nil } // only allowed one FREQ
                frequency = parse(frequency: varval[1])
                guard frequency != nil else { return nil } // invalid FREQ value
            case "COUNT":
                guard recurrenceEnd == nil else { return nil } // only one of either COUNT or UNTIL, not both
                recurrenceEnd = parse(count: varval[1])
                guard recurrenceEnd != nil else { return nil } // invalid COUNT
            case "UNTIL":
                guard recurrenceEnd == nil else { return nil } // only one of either COUNT or UNTIL, not both
                recurrenceEnd = parse(until: varval[1])
                guard recurrenceEnd != nil else { return nil } // invalid UNTIL
            case "INTERVAL":
                guard interval == nil else { return nil } // only allowed one INTERVAL
                interval = parse(interval: varval[1])
                guard interval != nil else { return nil } // invalid INTERVAL
            case "BYMONTH":
                guard monthsOfTheYear == nil else { return nil } // only allowed one BYMONTH
                monthsOfTheYear = parse(byMonth: varval[1])
                guard monthsOfTheYear != nil else { return nil } // invalid BYMONTH
            case "BYDAY":
                guard daysOfTheWeek == nil else { return nil } // only allowed one BYDAY
                daysOfTheWeek = parse(byDay: varval[1])
                guard daysOfTheWeek != nil else { return nil } // invalid BYDAY
            case "WKST":
                guard firstDayOfTheWeek == nil else { return nil } // only allowed one WKST
                firstDayOfTheWeek = parse(byWeekStart: varval[1])
                guard firstDayOfTheWeek != nil else { return nil } // invalid WKST
            case "BYMONTHDAY":
                guard daysOfTheMonth == nil else { return nil } // only allowed one BYMONTHDAY
                daysOfTheMonth = parse(byMonthDay: varval[1])
                guard daysOfTheMonth != nil else { return nil } // invalid BYMONTHDAY
            case "BYYEARDAY":
                guard daysOfTheYear == nil else { return nil } // only allowed one BYYEARDAY
                daysOfTheYear = parse(byYearDay: varval[1])
                guard daysOfTheYear != nil else { return nil } // invalid BYYEARDAY
            case "BYWEEKNO":
                guard weeksOfTheYear == nil else { return nil } // only allowed one BYWEEKNO
                weeksOfTheYear = parse(byWeekNo: varval[1])
                guard weeksOfTheYear != nil else { return nil } // invalid BYWEEKNO
            case "BYSETPOS":
                guard setPositions == nil else { return nil } // only allowed one BYSETPOS
                setPositions = parse(bySetPosition: varval[1])
                guard setPositions != nil else { return nil } // invalid BYSETPOS
                /* Not supported by EKRecurrenceRule
            case "BYHOUR":
                return nil
            case "BYMINUTE":
                return nil
            case "BYSECOND":
                return nil
                 */
            default:
                return nil
            }
        }

        if let frequency = frequency {
            return RWMRecurrenceRule(recurrenceWith: frequency, interval: interval, daysOfTheWeek: daysOfTheWeek, daysOfTheMonth: daysOfTheMonth, monthsOfTheYear: monthsOfTheYear, weeksOfTheYear: weeksOfTheYear, daysOfTheYear: daysOfTheYear, setPositions: setPositions, end: recurrenceEnd, firstDay: firstDayOfTheWeek)
        } else {
            return nil // no FREQ
        }
    }

    /// Returns the RRULE string represented by the provided recurrence rule.
    ///
    /// - Parameter from: The recurrence rule.
    /// - Returns: The RRULE string.
    public func rule(from: RWMRecurrenceRule) -> String {
        var parts = [String]()
        parts.append("FREQ=\(string(from: from.frequency))")

        if let interval = from.interval {
            parts.append("INTERVAL=\(interval)")
        }
        if let end = from.recurrenceEnd {
            if let date = end.endDate {
                parts.append("UNTIL=\(untilFormat.string(from: date))")
            } else {
                parts.append("COUNT=\(end.count)")
            }
        }
        if let wkst = from.firstDayOfTheWeek {
            parts.append("WKST=\(string(from: wkst))")
        }
        if let nums = from.weeksOfTheYear {
            parts.append("BYWEEKNO=\(string(from: nums))")
        }
        if let days = from.daysOfTheWeek {
            parts.append("BYDAY=\(string(from: days))")
        }
        if let nums = from.monthsOfTheYear {
            parts.append("BYMONTH=\(string(from: nums))")
        }
        if let nums = from.daysOfTheMonth {
            parts.append("BYMONTHDAY=\(string(from: nums))")
        }
        if let nums = from.daysOfTheYear {
            parts.append("BYYEARDAY=\(string(from: nums))")
        }
        if let nums = from.setPositions {
            parts.append("BYSETPOS=\(string(from: nums))")
        }

        return "RRULE:" + parts.joined(separator: ";")
    }

    private func parse(frequency: String) -> RWMRecurrenceFrequency? {
        switch frequency {
        case "DAILY":
            return .daily
        case "WEEKLY":
            return .weekly
        case "MONTHLY":
            return .monthly
        case "YEARLY":
            return .yearly
        case "HOURLY":
            return nil // not supported by EKRecurrenceRule
        case "MINUTELY":
            return nil // not supported by EKRecurrenceRule
        case "SECONDLY":
            return nil // not supported by EKRecurrenceRule
        default:
            return nil
        }
    }

    private func string(from: RWMRecurrenceFrequency) -> String {
        switch from {
        case .daily:
            return "DAILY"
        case .weekly:
            return "WEEKLY"
        case .monthly:
            return "MONTHLY"
        case .yearly:
            return "YEARLY"
        }
    }

    private func parse(count: String) -> RWMRecurrenceEnd? {
        if let cnt = Int(count) {
            return RWMRecurrenceEnd(occurrenceCount: cnt)
        } else {
            return nil
        }
    }

    private func parse(until: String) -> RWMRecurrenceEnd? {
        let df = DateFormatter()
        df.locale = Locale(identifier: "en_US_POSIX")
        for format in [ "yyyyMMdd'T'HHmmssX", "yyyyMMdd'T'HHmmss", "'TZID'=VV:yyyyMMdd'T'HHmmss", "yyyyMMdd" ] {
            df.dateFormat = format
            if let date = df.date(from: until) {
                return RWMRecurrenceEnd(end: date)
            }
        }

        return nil
    }

    private func parse(interval: String) -> Int? {
        if let cnt = Int(interval) {
            return cnt
        } else {
            return nil
        }
    }

    private func parseNumberList(_ list: String) -> [Int]? {
        var res = [Int]()
        let parts = list.components(separatedBy: ",")
        for part in parts {
            if let num = Int(part) {
                res.append(num)
            } else {
                return nil
            }
        }

        return res
    }

    private func string(from: [Int]) -> String {
        return from.map { String($0) }.joined(separator: ",")
    }

    private func parse(byMonth: String) -> [Int]? {
        return parseNumberList(byMonth)
    }

    private func parse(byWeekStart: String) -> RWMWeekday? {
        switch byWeekStart {
        case "SU":
            return .sunday
        case "MO":
            return .monday
        case "TU":
            return .tuesday
        case "WE":
            return .wednesday
        case "TH":
            return .thursday
        case "FR":
            return .friday
        case "SA":
            return .saturday
        default:
            return nil
        }
    }

    private func string(from: RWMWeekday) -> String {
        switch from {
        case .sunday:
            return "SU"
        case .monday:
            return "MO"
        case .tuesday:
            return "TU"
        case .wednesday:
            return "WE"
        case .thursday:
            return "TH"
        case .friday:
            return "FR"
        case .saturday:
            return "SA"
        }
    }

    private func parse(byDay: String) -> [RWMRecurrenceDayOfWeek]? {
        var res = [RWMRecurrenceDayOfWeek]()
        let parts = byDay.components(separatedBy: ",")
        for part in parts {
            let scanner = Scanner(string: part)
            var count = 0
            scanner.scanInt(&count)
            var weekday: NSString?
            if scanner.scanCharacters(from: .alphanumerics, into: &weekday) && scanner.isAtEnd {
                if let weekday = weekday, let dow = parse(byWeekStart: weekday as String) {
                    let rec = count == 0 ? RWMRecurrenceDayOfWeek(dow) : RWMRecurrenceDayOfWeek(dow, weekNumber: count)
                    res.append(rec)
                } else {
                    return nil
                }
            } else {
                return nil
            }
        }

        return res
    }

    private func string(from: [RWMRecurrenceDayOfWeek]) -> String {
        return from.map {
            var res = ""
            if $0.weekNumber != 0 {
                res += String($0.weekNumber)
            }
            res += string(from: $0.dayOfTheWeek)
            return res
        }.joined(separator: ",")
    }

    private func parse(byMonthDay: String) -> [Int]? {
        return parseNumberList(byMonthDay)
    }

    private func parse(byYearDay: String) -> [Int]? {
        return parseNumberList(byYearDay)
    }

    private func parse(byWeekNo: String) -> [Int]? {
        return parseNumberList(byWeekNo)
    }

    private func parse(bySetPosition: String) -> [Int]? {
        return parseNumberList(bySetPosition)
    }
}

/// The RWMRecurrenceEnd struct defines the end of a recurrence rule defined by an RWMRecurrenceRule object.
/// The recurrence end can be specified by a date (date-based) or by a maximum count of occurrences (count-based).
/// An event which is set to never end should have its RWMRecurrenceEnd set to nil.
public struct RWMRecurrenceEnd: Codable, Equatable {
    /// The end date of the recurrence end, or `nil` if the recurrence end is count-based.
    public let endDate: Date?
    /// The occurrence count of the recurrence end, or `0` if the recurrence end is date-based.
    public let count: Int

    /// Initializes and returns a date-based recurrence end with a given end date.
    ///
    /// - Parameter end: The end date.
    public init(end: Date) {
        self.endDate = end
        self.count = 0
    }

    /// Initializes and returns a count-based recurrence end with a given maximum occurrence count.
    ///
    /// - Parameter occurrenceCount: The maximum occurrence count.
    public init(occurrenceCount: Int) {
        self.endDate = nil
        if occurrenceCount > 0 {
            self.count = occurrenceCount
        } else {
            fatalError("occurrenceCount must be > 0")
        }
    }

    public static func==(lhs: RWMRecurrenceEnd, rhs: RWMRecurrenceEnd) -> Bool {
        if let ldate = lhs.endDate {
            if let rdate = rhs.endDate {
                return ldate == rdate // both are dates
            } else {
                return false // one date and one count
            }
        } else {
            if rhs.endDate != nil {
                return false // one date and one count
            } else {
                return lhs.count == rhs.count // both are counts
            }
        }
    }
}

/// The `RWMRecurrenceDayOfWeek` struct represents a day of the week for use with an `RWMRecurrenceRule` object.
/// A day of the week can optionally have a week number, indicating a specific day in the recurrence rule’s frequency.
/// For example, a day of the week with a day value of `Tuesday` and a week number of `2` would represent the second
/// Tuesday of every month in a monthly recurrence rule, and the second Tuesday of every year in a yearly recurrence
/// rule. A day of the week with a week number of `0` ignores its week number.
public struct RWMRecurrenceDayOfWeek: Codable, Equatable {
    /// The day of the week.
    public let dayOfTheWeek: RWMWeekday
    /// The week number of the day of the week.
    ///
    /// Values range from `-53` to `53`. A negative value indicates a value from the end of the range. `0` indicates the week number is irrelevant.
    public let weekNumber: Int

    /// Initializes and returns a day of the week with a given day and week number.
    ///
    /// - Parameters:
    ///   - dayOfTheWeek: The day of the week.
    ///   - weekNumber: The week number.
    public init(dayOfTheWeek: RWMWeekday, weekNumber: Int) {
        self.dayOfTheWeek = dayOfTheWeek
        if weekNumber < -53 || weekNumber > 53 {
            fatalError("weekNumber must be -53 to 53")
        } else {
            self.weekNumber = weekNumber
        }
    }

    /// Creates and returns a day of the week with a given day.
    ///
    /// - Parameter dayOfTheWeek: The day of the week.
    public init(_ dayOfTheWeek: RWMWeekday) {
        self.init(dayOfTheWeek: dayOfTheWeek, weekNumber: 0)
    }

    /// Creates and returns an autoreleased day of the week with a given day and week number.
    ///
    /// - Parameters:
    ///   - dayOfTheWeek: The day of the week.
    ///   - weekNumber: The week number.
    public init(_ dayOfTheWeek: RWMWeekday, weekNumber: Int) {
        self.init(dayOfTheWeek: dayOfTheWeek, weekNumber: weekNumber)
    }

    public static func==(lhs: RWMRecurrenceDayOfWeek, rhs: RWMRecurrenceDayOfWeek) -> Bool {
        return lhs.dayOfTheWeek == rhs.dayOfTheWeek && lhs.weekNumber == rhs.weekNumber
    }
}

/// The `RWMRecurrenceRule` class is used to describe the recurrence pattern for a recurring event.
public struct RWMRecurrenceRule: Codable, Equatable {
    /// The frequency of the recurrence rule.
    let frequency: RWMRecurrenceFrequency
    /// Specifies how often the recurrence rule repeats over the unit of time indicated by its frequency. For example, a recurrence rule with a frequency type of `.weekly` and an interval of `2` repeats every two weeks.
    let interval: Int?
    /// Indicates which day of the week the recurrence rule treats as the first day of the week. No value indicates that this property is not set for the recurrence rule.
    let firstDayOfTheWeek: RWMWeekday?
    /// The days of the week associated with the recurrence rule, as an array of `RWMRecurrenceDayOfWeek` objects.
    let daysOfTheWeek: [RWMRecurrenceDayOfWeek]?
    /// The days of the month associated with the recurrence rule, as an array of `Int`. Values can be from 1 to 31 and from -1 to -31. This property value is invalid with a frequency type of `.weekly`.
    let daysOfTheMonth: [Int]?
    /// The days of the year associated with the recurrence rule, as an array of `Int`. Values can be from 1 to 366 and from -1 to -366. This property value is valid only for recurrence rules initialized with a frequency type of `.yearly`.
    let daysOfTheYear: [Int]?
    /// The weeks of the year associated with the recurrence rule, as an array of `Int` objects. Values can be from 1 to 53 and from -1 to -53. This property value is valid only for recurrence rules initialized with specific weeks of the year and a frequency type of `.yearly`.
    let weeksOfTheYear: [Int]?
    /// The months of the year associated with the recurrence rule, as an array of `Int` objects. Values can be from 1 to 12. This property value is valid only for recurrence rules initialized with specific months of the year and a frequency type of `.yearly`.
    let monthsOfTheYear: [Int]?
    /// An array of ordinal numbers that filters which recurrences to include in the recurrence rule’s frequency. For example, a yearly recurrence rule that has a daysOfTheWeek value that specifies Monday through Friday, and a setPositions array containing 2 and -1, occurs only on the second weekday and last weekday of every year.
    let setPositions: [Int]?
    /// Indicates when the recurrence rule ends. This can be represented by an end date or a number of occurrences.
    let recurrenceEnd: RWMRecurrenceEnd?

    /// Initializes and returns a simple recurrence rule with a given frequency, interval, and end.
    ///
    /// - Parameters:
    ///   - type: Initializes and returns a simple recurrence rule with a given frequency, interval, and end.
    ///   - interval: The interval between instances of this recurrence. For example, a weekly recurrence rule with an interval of `2` occurs every other week. Must be greater than `0`.
    ///   - end: The end of the recurrence rule.
    public init?(recurrenceWith type: RWMRecurrenceFrequency, interval: Int?, end: RWMRecurrenceEnd?) {
        self.init(recurrenceWith: type, interval: interval, daysOfTheWeek: nil, daysOfTheMonth: nil, monthsOfTheYear: nil, weeksOfTheYear: nil, daysOfTheYear: nil, setPositions: nil, end: end, firstDay: nil)
    }

    /// Initializes and returns a recurrence rule with a given frequency and additional scheduling information.
    ///
    /// Returns `nil` is any invalid parameters are provided.
    ///
    /// Negative value indicate counting backwards from the end of the recurrence rule's frequency.
    ///
    /// - Parameters:
    ///   - type: The frequency of the recurrence rule. Can be daily, weekly, monthly, or yearly.
    ///   - interval: The interval between instances of this recurrence. For example, a weekly recurrence rule with an interval of `2` occurs every other week. Must be greater than `0`.
    ///   - days: The days of the week that the event occurs, as an array of `RWMRecurrenceDayOfWeek` objects.
    ///   - monthDays: The days of the month that the event occurs, as an array of `Int`. Values can be from 1 to 31 and from -1 to -31. This parameter is not valid for recurrence rules of type `.weekly`.
    ///   - months: The months of the year that the event occurs, as an array of `Int`. Values can be from 1 to 12.
    ///   - weeksOfTheYear: The weeks of the year that the event occurs, as an array of `Int`. Values can be from 1 to 53 and from -1 to -53. This parameter is only valid for recurrence rules of type `.yearly`.
    ///   - daysOfTheYear: The days of the year that the event occurs, as an array of `Int`. Values can be from 1 to 366 and from -1 to -366. This parameter is only valid for recurrence rules of type `.yearly`.
    ///   - setPositions: An array of ordinal numbers that filters which recurrences to include in the recurrence rule’s frequency. See `setPositions` for more information.
    ///   - end: The end of the recurrence rule.
    ///   - firstDay: Indicates what day of the week to be used as the first day of a week. Defaults to Monday.
    public init?(recurrenceWith type: RWMRecurrenceFrequency, interval: Int?, daysOfTheWeek days: [RWMRecurrenceDayOfWeek]?, daysOfTheMonth monthDays: [Int]?, monthsOfTheYear months: [Int]?, weeksOfTheYear: [Int]?, daysOfTheYear: [Int]?, setPositions: [Int]?, end: RWMRecurrenceEnd?, firstDay: RWMWeekday?) {
        // NOTE - See https://icalendar.org/iCalendar-RFC-5545/3-3-10-recurrence-rule.html

        if let interval = interval, interval <= 0 { return nil } // If specified, INTERVAL must be 1 or more
        if let days = days {
            // In daily or weekly mode or in yearly mode with week numbers, the days should not have a week number.
            if (type != .monthly && type != .yearly) || (type == .yearly && weeksOfTheYear != nil) {
                for day in days {
                    if day.weekNumber != 0 { return nil }
                }
            }
        }
        if let daysOfMonth = monthDays {
            guard type != .weekly else { return nil }

            for day in daysOfMonth {
                if day < -31 || day > 31 || day == 0 { return nil }
            }
        }
        if let monthsOfYear = months {
            for month in monthsOfYear {
                if month < 1 || month > 12 { return nil }
            }
        }
        if let weeksOfTheYear = weeksOfTheYear {
            guard type == .yearly else { return nil }

            for week in weeksOfTheYear {
                if week < -53 || week > 53 || week == 0 { return nil }
            }
        }
        if let daysOfTheYear = daysOfTheYear {
            // Also supported by secondly, minutely, and hourly
            guard type == .yearly else { return nil }

            for day in daysOfTheYear {
                if day < -366 || day > 366 || day == 0 { return nil }
            }
        }
        if let setPositions = setPositions {
            for pos in setPositions {
                if pos < -366 || pos > 366 || pos == 0 { return nil }
            }
        }

        self.frequency = type
        self.interval = interval
        self.firstDayOfTheWeek = firstDay
        self.daysOfTheWeek = days
        self.daysOfTheMonth = monthDays
        self.daysOfTheYear = daysOfTheYear
        self.weeksOfTheYear = weeksOfTheYear
        self.monthsOfTheYear = months
        self.setPositions = setPositions
        self.recurrenceEnd = end
    }

    public init?(recurrenceWith rule: EKRecurrenceRule) {
            var daysOfTheWeek: [RWMRecurrenceDayOfWeek]?
            if let dows = rule.daysOfTheWeek {
                daysOfTheWeek = []
                for dow in dows {
                    if let rwmwd = RWMWeekday(rawValue: dow.dayOfTheWeek.rawValue) {
                        daysOfTheWeek?.append(RWMRecurrenceDayOfWeek(dayOfTheWeek: rwmwd, weekNumber: dow.weekNumber))
                    } else {
                        return nil
                    }
                }
            }

            let end: RWMRecurrenceEnd?
            if let rend = rule.recurrenceEnd {
                if let date = rend.endDate {
                    end = RWMRecurrenceEnd(end: date)
                } else {
                    end = RWMRecurrenceEnd(occurrenceCount: rend.occurrenceCount)
                }
            } else {
                end = nil
            }

            if let frequency = RWMRecurrenceFrequency(rawValue: rule.frequency.rawValue) {
                // For weekly recurrence rules with days of the week set, set the rule's firstDay if the current calendar
                // starts its week on a day other than Monday.
                var firstDay: RWMWeekday? = nil
                if frequency == .weekly && daysOfTheWeek != nil && Calendar.current.firstWeekday != 2 {
                    firstDay = RWMWeekday(rawValue: Calendar.current.firstWeekday)
                }

                self.init(recurrenceWith: frequency, interval: rule.interval == 1 ? nil : rule.interval, daysOfTheWeek: daysOfTheWeek, daysOfTheMonth: rule.daysOfTheMonth as! [Int]?, monthsOfTheYear: rule.monthsOfTheYear as! [Int]?, weeksOfTheYear: rule.weeksOfTheYear as! [Int]?, daysOfTheYear: rule.daysOfTheYear as! [Int]?, setPositions: rule.setPositions as! [Int]?, end: end, firstDay: firstDay)
            } else {
                return nil
            }
        }

    public static func==(lhs: RWMRecurrenceRule, rhs: RWMRecurrenceRule) -> Bool {
        return
            lhs.frequency == rhs.frequency &&
            lhs.interval == rhs.interval &&
            lhs.firstDayOfTheWeek == rhs.firstDayOfTheWeek &&
            lhs.daysOfTheWeek == rhs.daysOfTheWeek &&
            lhs.daysOfTheMonth == rhs.daysOfTheMonth &&
            lhs.daysOfTheYear == rhs.daysOfTheYear &&
            lhs.weeksOfTheYear == rhs.weeksOfTheYear &&
            lhs.monthsOfTheYear == rhs.monthsOfTheYear &&
            lhs.setPositions == rhs.setPositions &&
            lhs.recurrenceEnd == rhs.recurrenceEnd
    }
}

extension EKRecurrenceRule {
    /// This convenience initializer allows you to create an EKRecurrenceRule from a standard iCalendar RRULE
    /// string. Please see https://icalendar.org/iCalendar-RFC-5545/3-3-10-recurrence-rule.html for a reference
    /// to the RRULE syntax.
    /// Only frequencies of DAILY, WEEKLY, MONTHLY, and YEARLY are supported. Also note that there are many valid
    /// RRULE strings that will parse but EventKit may not process correctly.
    ///
    /// If `rrule` is an invalid RRULE, the result is `nil`.
    ///
    /// See `RWMRecurrenceRule isEventKitSafe` for details about RRULE values safe to be used with Event Kit.
    ///
    /// - Parameter rrule: The RRULE string in the form RRULE:FREQUENCY=...
    public convenience init?(recurrenceWith rrule: String) {
        if let rule = RWMRuleParser().parse(rule: rrule) {
            self.init(recurrenceWith: rule)
        } else {
            return nil
        }
    }

    /// Creates a new EKRecurrenceRule from a RWMRecurrenceRule. If `rule` can't be converted, the result is `nil`.
    ///
    /// Note that Event Kit may not properly process some recurrence rules.
    ///
    /// - Parameter rule: The RWMRecurrenceRule.
    public convenience init?(recurrenceWith rule: RWMRecurrenceRule) {
        var daysOfTheWeek: [EKRecurrenceDayOfWeek]?
        if let dows = rule.daysOfTheWeek {
            daysOfTheWeek = []
            for dow in dows {
                if let ekwd = EKWeekday(rawValue: dow.dayOfTheWeek.rawValue) {
                    daysOfTheWeek?.append(EKRecurrenceDayOfWeek(dayOfTheWeek: ekwd, weekNumber: dow.weekNumber))
                } else {
                    return nil
                }
            }
        }

        let end: EKRecurrenceEnd?
        if let rend = rule.recurrenceEnd {
            if let date = rend.endDate {
                end = EKRecurrenceEnd(end: date)
            } else {
                end = EKRecurrenceEnd(occurrenceCount: rend.count)
            }
        } else {
            end = nil
        }

        if let frequency = EKRecurrenceFrequency(rawValue: rule.frequency.rawValue) {
            self.init(recurrenceWith: frequency, interval: rule.interval ?? 1, daysOfTheWeek: daysOfTheWeek, daysOfTheMonth: rule.daysOfTheMonth as [NSNumber]?, monthsOfTheYear: rule.monthsOfTheYear as [NSNumber]?, weeksOfTheYear: rule.weeksOfTheYear as [NSNumber]?, daysOfTheYear: rule.daysOfTheYear as [NSNumber]?, setPositions: rule.setPositions as [NSNumber]?, end: end)
        } else {
            return nil
        }
    }

    /// Returns the RRULE representation. If the sender can't be processed, the result is `nil`.
    public var rrule: String? {
        if let rule = RWMRecurrenceRule(recurrenceWith: self) {
            let parser = RWMRuleParser()

            return parser.rule(from: rule)
        } else {
            return nil
        }
    }
}

extension Calendar {
    /// Returns the range of the given weekday for the supplied year or month of year.
    ///
    ///  Examples:
    ///    - To find out how many Tuesdays there are in 2018, pass in `3` for the `weekday` and `2018` for the `year` and the default of `0` for the `month`.
    ///    - To find out how many Saturdays there are in May of 2018, pass in `7` for the `weekday`, `2018` for the `year`, and `5` for the `month`.
    ///    - To find out how many Mondays there are in the last month of 2018, pass in `2` for the `weekday`, `2018` for the `year`, and `-1` for the `month`.
    ///
    /// - Parameters:
    ///   - weekday: The day of the week. Values range from 1 to 7, with Sunday being 1.
    ///   - year: A calendar year.
    ///   - month: A month within the calendar year. The value of `0` means the month is ignored. Negative values start from the last month of the year. `-1` is the last month. `-2` is the next-to-last month, etc.
    /// - Returns: A range from `1` through `n` where `n` is the number of times the given weekday appears in the year or month of the year. If `month` is out of range for the year, the result is `nil`.
    public func range(of weekday: Int, in year: Int, month: Int = 0) -> ClosedRange<Int>? {
        if month > 0 {
            let comps = DateComponents(year: year, month: month, weekday: weekday, weekdayOrdinal: -1)
            if let date = self.date(from: comps) {
                let count = self.component(.weekdayOrdinal, from: date)

                return 1...count
            }
        } else {
            // Get first day of year for the given weekday
            let startComps = DateComponents(year: year, month: 1, weekday: weekday, weekdayOrdinal: 1)
            // Get last day of year for the given weekday
            let finishComps = DateComponents(year: year, month: 12, weekday: weekday, weekdayOrdinal: -1)
            if let startDate = self.date(from: startComps), let finishDate = self.date(from: finishComps) {
                // Get the number of days between the two dates
                let days = self.dateComponents([.day], from: startDate, to: finishDate).day!

                return 1...(days / 7 + 1)
            }
        }

        return nil
    }

    /// Converts relative components to normalized components.
    ///
    /// The following relative components are normalized:
    ///   - year set, month set, weekday set, weekday ordinal set to +/- instance of weekday within month
    ///   - year set, no month, weekday set, weekday ordinal set to +/- instance of weekday within year
    ///   - year set, month set, day set to +/- day of month
    ///   - year set, no month, day set to +/- day of year
    ///
    /// All other components are returned as-is.
    ///
    /// - Parameter components: The relative date components.
    /// - Returns: The normalized date components.
    func relativeComponents(from components: DateComponents) -> DateComponents {
        var newComponents = components

        if let year = components.year {
            if let weekday = components.weekday, let ordinal = components.weekdayOrdinal {
                if ordinal < 0 {
                    if let month = components.month {
                        if let rng = self.range(of: weekday, in: year, month: month) {
                            newComponents.weekdayOrdinal = rng.count + ordinal + 1
                        }
                    } else {
                        if let rng = self.range(of: weekday, in: year) {
                            newComponents.weekdayOrdinal = rng.count + ordinal + 1
                        }
                    }
                } else {
                    // Calendar already handles positive weekdayOrdinal
                }
            } else if let day = components.day {
                if components.weekday == nil {
                    if let month = components.month {
                        if day < 0 {
                            if let startOfMonth = self.date(from: DateComponents(year: year, month: month, day: 1)),
                                let daysInMonth = self.range(of: .day, in: .month, for: startOfMonth)?.count {
                                newComponents.day = daysInMonth + day + 1
                            }
                        } else {
                            // Calendar already handles positive day
                        }
                    } else {
                        if day < 0 {
                            if let startOfYear = self.date(from: DateComponents(year: year, month: 1, day: 1)),
                                let daysInYear = self.range(of: .day, in: .year, for: startOfYear)?.count {
                                newComponents.day = daysInYear + day + 1
                            }
                        } else {
                            // Calendar already handles positive day
                        }
                    }
                }
            }
        }

        return newComponents
    }

    public func date(fromRelative components: DateComponents) -> Date? {
        return self.date(from: self.relativeComponents(from: components))
    }

    public func date(_ date: Date, matchesRelativeComponents components: DateComponents) -> Bool {
        return self.date(date, matchesComponents: self.relativeComponents(from: components))
    }
}

public class SwiftDeviceCalendarPlugin: NSObject, FlutterPlugin {
    struct DeviceCalendar: Codable {
        let id: String
        let name: String
        let isReadOnly: Bool
        let isDefault: Bool
        let color: Int
        let accountName: String
        let accountType: String
    }
    
    struct Event: Codable {
        let eventId: String
        let calendarId: String
        let title: String
        let description: String?
        let start: Int64
        let end: Int64
        let startTimeZone: String?
        let allDay: Bool
        let attendees: [Attendee]
        let location: String?
        let url: String?
        let recurrenceRule: String?
        let organizer: Attendee?
        let reminders: [Reminder]
        let availability: Availability?
    }





    struct Attendee: Codable {
        let name: String?
        let emailAddress: String
        let role: Int
        let attendanceStatus: Int
    }
    
    struct Reminder: Codable {
        let minutes: Int
    }
    
    enum Availability: String, Codable {
        case BUSY
		case FREE
		case TENTATIVE
		case UNAVAILABLE
    }



    static let channelName = "plugins.builttoroam.com/device_calendar"
    let notFoundErrorCode = "404"
    let notAllowed = "405"
    let genericError = "500"
    let unauthorizedErrorCode = "401"
    let unauthorizedErrorMessage = "The user has not allowed this application to modify their calendar(s)"
    let calendarNotFoundErrorMessageFormat = "The calendar with the ID %@ could not be found"
    let calendarReadOnlyErrorMessageFormat = "Calendar with ID %@ is read-only"
    let eventNotFoundErrorMessageFormat = "The event with the ID %@ could not be found"
    let eventStore = EKEventStore()
    let requestPermissionsMethod = "requestPermissions"
    let hasPermissionsMethod = "hasPermissions"
    let retrieveCalendarsMethod = "retrieveCalendars"
    let retrieveEventsMethod = "retrieveEvents"
    let retrieveSourcesMethod = "retrieveSources"
    let createOrUpdateEventMethod = "createOrUpdateEvent"
    let createCalendarMethod = "createCalendar"
    let deleteCalendarMethod = "deleteCalendar"
    let deleteEventMethod = "deleteEvent"
    let deleteEventInstanceMethod = "deleteEventInstance"
    let calendarIdArgument = "calendarId"
    let startDateArgument = "startDate"
    let endDateArgument = "endDate"
    let eventIdArgument = "eventId"
    let eventIdsArgument = "eventIds"
    let eventTitleArgument = "eventTitle"
    let eventDescriptionArgument = "eventDescription"
    let eventAllDayArgument = "eventAllDay"
    let eventStartDateArgument =  "eventStartDate"
    let eventEndDateArgument = "eventEndDate"
    let eventStartTimeZoneArgument = "eventStartTimeZone"
    let eventLocationArgument = "eventLocation"
    let eventURLArgument = "eventURL"
    let attendeesArgument = "attendees"
    let recurrenceRuleArgument = "recurrenceRule"
    let recurrenceFrequencyArgument = "recurrenceFrequency"
    let totalOccurrencesArgument = "totalOccurrences"
    let intervalArgument = "interval"
    let daysOfWeekArgument = "daysOfWeek"
    let dayOfMonthArgument = "dayOfMonth"
    let monthOfYearArgument = "monthOfYear"
    let weekOfMonthArgument = "weekOfMonth"
    let nameArgument = "name"
    let emailAddressArgument = "emailAddress"
    let roleArgument = "role"
    let remindersArgument = "reminders"
    let minutesArgument = "minutes"
    let followingInstancesArgument = "followingInstances"
    let calendarNameArgument = "calendarName"
    let calendarColorArgument = "calendarColor"
    let availabilityArgument = "availability"
    let validFrequencyTypes = [EKRecurrenceFrequency.daily, EKRecurrenceFrequency.weekly, EKRecurrenceFrequency.monthly, EKRecurrenceFrequency.yearly]
    
    public static func register(with registrar: FlutterPluginRegistrar) {
        let channel = FlutterMethodChannel(name: channelName, binaryMessenger: registrar.messenger())
        let instance = SwiftDeviceCalendarPlugin()
        registrar.addMethodCallDelegate(instance, channel: channel)
    }
    
    public func handle(_ call: FlutterMethodCall, result: @escaping FlutterResult) {
        switch call.method {
        case requestPermissionsMethod:
            requestPermissions(result)
        case hasPermissionsMethod:
            hasPermissions(result)
        case retrieveCalendarsMethod:
            retrieveCalendars(result)
        case retrieveEventsMethod:
            retrieveEvents(call, result)
        case createOrUpdateEventMethod:
            createOrUpdateEvent(call, result)
        case deleteEventMethod:
            deleteEvent(call, result)
        case deleteEventInstanceMethod:
            deleteEvent(call, result)
        case createCalendarMethod:
            createCalendar(call, result)
        case deleteCalendarMethod:
            deleteCalendar(call, result)
        default:
            result(FlutterMethodNotImplemented)
        }
    }
    
    private func hasPermissions(_ result: FlutterResult) {
        let hasPermissions = hasEventPermissions()
        result(hasPermissions)
    }
    
    private func createCalendar(_ call: FlutterMethodCall, _ result: FlutterResult) {
        let arguments = call.arguments as! Dictionary<String, AnyObject>
        let calendar = EKCalendar.init(for: EKEntityType.event, eventStore: eventStore)
        do {
            calendar.title = arguments[calendarNameArgument] as! String
            let calendarColor = arguments[calendarColorArgument] as? String
            
            if (calendarColor != nil) {
                calendar.cgColor = UIColor(hex: calendarColor!)?.cgColor
            }
            else {
                calendar.cgColor = UIColor(red: 255, green: 0, blue: 0, alpha: 0).cgColor // Red colour as a default
            }
            
            let localSources = eventStore.sources.filter { $0.sourceType == .local }
            
            if (!localSources.isEmpty) {
                calendar.source = localSources.first
                
                try eventStore.saveCalendar(calendar, commit: true)
                result(calendar.calendarIdentifier)
            }
            else {
                result(FlutterError(code: self.genericError, message: "Local calendar was not found.", details: nil))
            }
        }
        catch {
            eventStore.reset()
            result(FlutterError(code: self.genericError, message: error.localizedDescription, details: nil))
        }
    }
    
    private func retrieveCalendars(_ result: @escaping FlutterResult) {
        checkPermissionsThenExecute(permissionsGrantedAction: {
            let ekCalendars = self.eventStore.calendars(for: .event)
            let defaultCalendar = self.eventStore.defaultCalendarForNewEvents
            var calendars = [DeviceCalendar]()
            for ekCalendar in ekCalendars {
                let calendar = DeviceCalendar(
                    id: ekCalendar.calendarIdentifier,
                    name: ekCalendar.title,
                    isReadOnly: !ekCalendar.allowsContentModifications,
                    isDefault: defaultCalendar?.calendarIdentifier == ekCalendar.calendarIdentifier,
                    color: UIColor(cgColor: ekCalendar.cgColor).rgb()!,
                    accountName: ekCalendar.source.title,
                    accountType: getAccountType(ekCalendar.source.sourceType))
                calendars.append(calendar)
            }
            
            self.encodeJsonAndFinish(codable: calendars, result: result)
        }, result: result)
    }
    
    private func deleteCalendar(_ call: FlutterMethodCall, _ result: @escaping FlutterResult) {
        checkPermissionsThenExecute(permissionsGrantedAction: {
            let arguments = call.arguments as! Dictionary<String, AnyObject>
            let calendarId = arguments[calendarIdArgument] as! String
            
            let ekCalendar = self.eventStore.calendar(withIdentifier: calendarId)
            if ekCalendar == nil {
                self.finishWithCalendarNotFoundError(result: result, calendarId: calendarId)
                return
            }
            
            if !(ekCalendar!.allowsContentModifications) {
                self.finishWithCalendarReadOnlyError(result: result, calendarId: calendarId)
                return
            }
            
            do {
                try self.eventStore.removeCalendar(ekCalendar!, commit: true)
                result(true)
            } catch {
                self.eventStore.reset()
                result(FlutterError(code: self.genericError, message: error.localizedDescription, details: nil))
            }
        }, result: result)
    }
    
    private func getAccountType(_ sourceType: EKSourceType) -> String {
        switch (sourceType) {
        case .local:
            return "Local";
        case .exchange:
            return "Exchange";
        case .calDAV:
            return "CalDAV";
        case .mobileMe:
            return "MobileMe";
        case .subscribed:
            return "Subscribed";
        case .birthdays:
            return "Birthdays";
        default:
            return "Unknown";
        }
    }
    
    private func retrieveEvents(_ call: FlutterMethodCall, _ result: @escaping FlutterResult) {
        checkPermissionsThenExecute(permissionsGrantedAction: {
            let arguments = call.arguments as! Dictionary<String, AnyObject>
            let calendarId = arguments[calendarIdArgument] as! String
            let startDateMillisecondsSinceEpoch = arguments[startDateArgument] as? NSNumber
            let endDateDateMillisecondsSinceEpoch = arguments[endDateArgument] as? NSNumber
            let eventIds = arguments[eventIdsArgument] as? [String]
            var events = [Event]()
            let specifiedStartEndDates = startDateMillisecondsSinceEpoch != nil && endDateDateMillisecondsSinceEpoch != nil
            if specifiedStartEndDates {
                let startDate = Date (timeIntervalSince1970: startDateMillisecondsSinceEpoch!.doubleValue / 1000.0)
                let endDate = Date (timeIntervalSince1970: endDateDateMillisecondsSinceEpoch!.doubleValue / 1000.0)
                let ekCalendar = self.eventStore.calendar(withIdentifier: calendarId)
                let predicate = self.eventStore.predicateForEvents(withStart: startDate, end: endDate, calendars: [ekCalendar!])
                let ekEvents = self.eventStore.events(matching: predicate)
                for ekEvent in ekEvents {
                    let event = createEventFromEkEvent(calendarId: calendarId, ekEvent: ekEvent)
                    events.append(event)
                }
            }
            
            if eventIds == nil {
                self.encodeJsonAndFinish(codable: events, result: result)
                return
            }
            
            if specifiedStartEndDates {
                events = events.filter({ (e) -> Bool in
                    e.calendarId == calendarId && eventIds!.contains(e.eventId)
                })
                
                self.encodeJsonAndFinish(codable: events, result: result)
                return
            }
            
            for eventId in eventIds! {
                let ekEvent = self.eventStore.event(withIdentifier: eventId)
                if ekEvent == nil {
                    continue
                }
                
                let event = createEventFromEkEvent(calendarId: calendarId, ekEvent: ekEvent!)
                events.append(event)
            }
            
            self.encodeJsonAndFinish(codable: events, result: result)
        }, result: result)
    }
    
    private func createEventFromEkEvent(calendarId: String, ekEvent: EKEvent) -> Event {
        var attendees = [Attendee]()
        if ekEvent.attendees != nil {
            for ekParticipant in ekEvent.attendees! {
                let attendee = convertEkParticipantToAttendee(ekParticipant: ekParticipant)
                if attendee == nil {
                    continue
                }
                
                attendees.append(attendee!)
            }
        }
        
        var reminders = [Reminder]()
        if ekEvent.alarms != nil {
            for alarm in ekEvent.alarms! {
                reminders.append(Reminder(minutes: Int(-alarm.relativeOffset / 60)))
            }
        }
        
        let recurrenceRule = parseEKRecurrenceRules(ekEvent)

        var rruleString : String?

        if (recurrenceRule != nil) {
            let parser = RWMRuleParser()

            rruleString = parser.rule(from: recurrenceRule!)
        }



        let event = Event(
            eventId: ekEvent.eventIdentifier,
            calendarId: calendarId,
            title: ekEvent.title ?? "New Event",
            description: ekEvent.notes,
            start: Int64(ekEvent.startDate.millisecondsSinceEpoch),
            end: Int64(ekEvent.endDate.millisecondsSinceEpoch),
            startTimeZone: ekEvent.timeZone?.identifier,
            allDay: ekEvent.isAllDay,
            attendees: attendees,
            location: ekEvent.location,
            url: ekEvent.url?.absoluteString,
//             recurrenceRule: recurrenceRule,
            recurrenceRule: rruleString,
            organizer: convertEkParticipantToAttendee(ekParticipant: ekEvent.organizer),
            reminders: reminders,
            availability: convertEkEventAvailability(ekEventAvailability: ekEvent.availability)
        )

        return event
    }
    
    private func convertEkParticipantToAttendee(ekParticipant: EKParticipant?) -> Attendee? {
        if ekParticipant == nil || ekParticipant?.emailAddress == nil {
            return nil
        }
        
        let attendee = Attendee(name: ekParticipant!.name, emailAddress:  ekParticipant!.emailAddress!, role: ekParticipant!.participantRole.rawValue, attendanceStatus: ekParticipant!.participantStatus.rawValue)
        return attendee
    }
    
    private func convertEkEventAvailability(ekEventAvailability: EKEventAvailability?) -> Availability? {
        switch ekEventAvailability {
        case .busy:
			return Availability.BUSY
        case .free:
            return Availability.FREE
		case .tentative:
			return Availability.TENTATIVE
		case .unavailable:
			return Availability.UNAVAILABLE
        default:
            return nil
        }
    }
    
    private func parseEKRecurrenceRules(_ ekEvent: EKEvent) -> RWMRecurrenceRule? {
        var recurrenceRule: RWMRecurrenceRule?
        if ekEvent.hasRecurrenceRules {
            let ekRecurrenceRule = ekEvent.recurrenceRules![0]
//             print("EKRecurrenceRule: \(ekRecurrenceRule)")
            recurrenceRule = RWMRecurrenceRule(recurrenceWith: ekRecurrenceRule)
        }
        return recurrenceRule
    }
    
    private func createEKRecurrenceRules(_ arguments: [String : AnyObject]) -> [EKRecurrenceRule]?{

        let recurrenceRuleArguments = arguments[recurrenceRuleArgument] as? String
//         print("RWMRecurrenceRule: \(recurrenceRuleArguments)")

        if recurrenceRuleArguments == nil {
            return nil
        }
        var ekRecurrenceRule = [EKRecurrenceRule]()

        ekRecurrenceRule.append(EKRecurrenceRule(recurrenceWith: recurrenceRuleArguments!)!)

        return  ekRecurrenceRule
    }
    
    private func setAttendees(_ arguments: [String : AnyObject], _ ekEvent: EKEvent?) {
        let attendeesArguments = arguments[attendeesArgument] as? [Dictionary<String, AnyObject>]
        if attendeesArguments == nil {
            return
        }
        
        var attendees = [EKParticipant]()
        for attendeeArguments in attendeesArguments! {
            let name = attendeeArguments[nameArgument] as! String
            let emailAddress = attendeeArguments[emailAddressArgument] as! String
            let role = attendeeArguments[roleArgument] as! Int
            
            if (ekEvent!.attendees != nil) {
                let existingAttendee = ekEvent!.attendees!.first { element in
                    return element.emailAddress == emailAddress
                }
                if existingAttendee != nil && ekEvent!.organizer?.emailAddress != existingAttendee?.emailAddress{
                    attendees.append(existingAttendee!)
                    continue
                }
            }
            
            let attendee = createParticipant(
                name: name,
                emailAddress: emailAddress,
                role: role)
            
            if (attendee == nil) {
                continue
            }
            
            attendees.append(attendee!)
        }
        
        ekEvent!.setValue(attendees, forKey: "attendees")
    }
    
    private func createReminders(_ arguments: [String : AnyObject]) -> [EKAlarm]?{
        let remindersArguments = arguments[remindersArgument] as? [Dictionary<String, AnyObject>]
        if remindersArguments == nil {
            return nil
        }
        
        var reminders = [EKAlarm]()
        for reminderArguments in remindersArguments! {
            let minutes = reminderArguments[minutesArgument] as! Int
            reminders.append(EKAlarm.init(relativeOffset: 60 * Double(-minutes)))
        }
        
        return reminders
    }
    
    private func setAvailability(_ arguments: [String : AnyObject]) -> EKEventAvailability? {
        guard let availabilityValue = arguments[availabilityArgument] as? String else { 
            return .unavailable 
        }

        switch availabilityValue.uppercased() {
        case Availability.BUSY.rawValue:
            return .busy
        case Availability.FREE.rawValue:
            return .free
		case Availability.TENTATIVE.rawValue:
        	return .tentative
        case Availability.UNAVAILABLE.rawValue:
            return .unavailable
        default:
            return nil
        }
    }
    
    private func createOrUpdateEvent(_ call: FlutterMethodCall, _ result: @escaping FlutterResult) {
        checkPermissionsThenExecute(permissionsGrantedAction: {
            let arguments = call.arguments as! Dictionary<String, AnyObject>
            let calendarId = arguments[calendarIdArgument] as! String
            let eventId = arguments[eventIdArgument] as? String
            let isAllDay = arguments[eventAllDayArgument] as! Bool
            let startDateMillisecondsSinceEpoch = arguments[eventStartDateArgument] as! NSNumber
            let endDateDateMillisecondsSinceEpoch = arguments[eventEndDateArgument] as! NSNumber
            let startDate = Date (timeIntervalSince1970: startDateMillisecondsSinceEpoch.doubleValue / 1000.0)
            let endDate = Date (timeIntervalSince1970: endDateDateMillisecondsSinceEpoch.doubleValue / 1000.0)
            let startTimeZoneString = arguments[eventStartTimeZoneArgument] as? String
            let title = arguments[self.eventTitleArgument] as! String
            let description = arguments[self.eventDescriptionArgument] as? String
            let location = arguments[self.eventLocationArgument] as? String
            let url = arguments[self.eventURLArgument] as? String
            let ekCalendar = self.eventStore.calendar(withIdentifier: calendarId)
            if (ekCalendar == nil) {
                self.finishWithCalendarNotFoundError(result: result, calendarId: calendarId)
                return
            }
            
            if !(ekCalendar!.allowsContentModifications) {
                self.finishWithCalendarReadOnlyError(result: result, calendarId: calendarId)
                return
            }
            
            var ekEvent: EKEvent?
            if eventId == nil {
                ekEvent = EKEvent.init(eventStore: self.eventStore)
            } else {
                ekEvent = self.eventStore.event(withIdentifier: eventId!)
                if(ekEvent == nil) {
                    self.finishWithEventNotFoundError(result: result, eventId: eventId!)
                    return
                }
            }
            
            ekEvent!.title = title
            ekEvent!.notes = description
            ekEvent!.isAllDay = isAllDay
            ekEvent!.startDate = startDate
            if (isAllDay) { ekEvent!.endDate = startDate }
            else {
                ekEvent!.endDate = endDate
                
                let timeZone = TimeZone(identifier: startTimeZoneString ?? TimeZone.current.identifier) ?? .current
                ekEvent!.timeZone = timeZone
            }
            ekEvent!.calendar = ekCalendar!
            ekEvent!.location = location

            // Create and add URL object only when if the input string is not empty or nil
            if let urlCheck = url, !urlCheck.isEmpty {
                let iosUrl = URL(string: url ?? "")
                ekEvent!.url = iosUrl
            }
            else {
                ekEvent!.url = nil
            }
            
            ekEvent!.recurrenceRules = createEKRecurrenceRules(arguments)
            setAttendees(arguments, ekEvent)
            ekEvent!.alarms = createReminders(arguments)
            
            if let availability = setAvailability(arguments) {
                ekEvent!.availability = availability
            }
            
            do {
                try self.eventStore.save(ekEvent!, span: .futureEvents)
                result(ekEvent!.eventIdentifier)
            } catch {
                self.eventStore.reset()
                result(FlutterError(code: self.genericError, message: error.localizedDescription, details: nil))
            }
        }, result: result)
    }
    
    private func createParticipant(name: String, emailAddress: String, role: Int) -> EKParticipant? {
        let ekAttendeeClass: AnyClass? = NSClassFromString("EKAttendee")
        if let type = ekAttendeeClass as? NSObject.Type {
            let participant = type.init()
            participant.setValue(name, forKey: "displayName")
            participant.setValue(emailAddress, forKey: "emailAddress")
            participant.setValue(role, forKey: "participantRole")
            return participant as? EKParticipant
        }
        return nil
    }
    
    private func deleteEvent(_ call: FlutterMethodCall, _ result: @escaping FlutterResult) {
        checkPermissionsThenExecute(permissionsGrantedAction: {
            let arguments = call.arguments as! Dictionary<String, AnyObject>
            let calendarId = arguments[calendarIdArgument] as! String
            let eventId = arguments[eventIdArgument] as! String
            let startDateNumber = arguments[eventStartDateArgument] as? NSNumber
            let endDateNumber = arguments[eventEndDateArgument] as? NSNumber
            let followingInstances = arguments[followingInstancesArgument] as? Bool
            
            let ekCalendar = self.eventStore.calendar(withIdentifier: calendarId)
            if ekCalendar == nil {
                self.finishWithCalendarNotFoundError(result: result, calendarId: calendarId)
                return
            }
            
            if !(ekCalendar!.allowsContentModifications) {
                self.finishWithCalendarReadOnlyError(result: result, calendarId: calendarId)
                return
            }
            
            if (startDateNumber == nil && endDateNumber == nil && followingInstances == nil) {
                let ekEvent = self.eventStore.event(withIdentifier: eventId)
                if ekEvent == nil {
                    self.finishWithEventNotFoundError(result: result, eventId: eventId)
                    return
                }
                
                do {
                    try self.eventStore.remove(ekEvent!, span: .futureEvents)
                    result(true)
                } catch {
                    self.eventStore.reset()
                    result(FlutterError(code: self.genericError, message: error.localizedDescription, details: nil))
                }
            }
            else {
                let startDate = Date (timeIntervalSince1970: startDateNumber!.doubleValue / 1000.0)
                let endDate = Date (timeIntervalSince1970: endDateNumber!.doubleValue / 1000.0)
                                
                let predicate = self.eventStore.predicateForEvents(withStart: startDate, end: endDate, calendars: nil)
                let foundEkEvents = self.eventStore.events(matching: predicate) as [EKEvent]?
                
                if foundEkEvents == nil || foundEkEvents?.count == 0 {
                    self.finishWithEventNotFoundError(result: result, eventId: eventId)
                    return
                }
                
                let ekEvent = foundEkEvents!.first(where: {$0.eventIdentifier == eventId})
                
                do {
                    if (!followingInstances!) {
                        try self.eventStore.remove(ekEvent!, span: .thisEvent, commit: true)
                    }
                    else {
                        try self.eventStore.remove(ekEvent!, span: .futureEvents, commit: true)
                    }
                    
                    result(true)
                } catch {
                    self.eventStore.reset()
                    result(FlutterError(code: self.genericError, message: error.localizedDescription, details: nil))
                }
            }
        }, result: result)
    }
    
    private func finishWithUnauthorizedError(result: @escaping FlutterResult) {
        result(FlutterError(code:self.unauthorizedErrorCode, message: self.unauthorizedErrorMessage, details: nil))
    }
    
    private func finishWithCalendarNotFoundError(result: @escaping FlutterResult, calendarId: String) {
        let errorMessage = String(format: self.calendarNotFoundErrorMessageFormat, calendarId)
        result(FlutterError(code:self.notFoundErrorCode, message: errorMessage, details: nil))
    }
    
    private func finishWithCalendarReadOnlyError(result: @escaping FlutterResult, calendarId: String) {
        let errorMessage = String(format: self.calendarReadOnlyErrorMessageFormat, calendarId)
        result(FlutterError(code:self.notAllowed, message: errorMessage, details: nil))
    }
    
    private func finishWithEventNotFoundError(result: @escaping FlutterResult, eventId: String) {
        let errorMessage = String(format: self.eventNotFoundErrorMessageFormat, eventId)
        result(FlutterError(code:self.notFoundErrorCode, message: errorMessage, details: nil))
    }
    
    private func encodeJsonAndFinish<T: Codable>(codable: T, result: @escaping FlutterResult) {
        do {
            let jsonEncoder = JSONEncoder()
            let jsonData = try jsonEncoder.encode(codable)
            let jsonString = String(data: jsonData, encoding: .utf8)
            result(jsonString)
        } catch {
            result(FlutterError(code: genericError, message: error.localizedDescription, details: nil))
        }
    }
    
    private func checkPermissionsThenExecute(permissionsGrantedAction: () -> Void, result: @escaping FlutterResult) {
        if hasEventPermissions() {
            permissionsGrantedAction()
            return
        }
        self.finishWithUnauthorizedError(result: result)
    }
    
    private func requestPermissions(completion: @escaping (Bool) -> Void) {
        if hasEventPermissions() {
            completion(true)
            return
        }
        eventStore.requestAccess(to: .event, completion: {
            (accessGranted: Bool, _: Error?) in
            completion(accessGranted)
        })
    }
    
    private func hasEventPermissions() -> Bool {
        let status = EKEventStore.authorizationStatus(for: .event)
        return status == EKAuthorizationStatus.authorized
    }
    
    private func requestPermissions(_ result: @escaping FlutterResult) {
        if hasEventPermissions()  {
            result(true)
        }
        eventStore.requestAccess(to: .event, completion: {
            (accessGranted: Bool, _: Error?) in
            result(accessGranted)
        })
    }
}

extension Date {
    func convert(from initTimeZone: TimeZone, to targetTimeZone: TimeZone) -> Date {
        let delta = TimeInterval(initTimeZone.secondsFromGMT() - targetTimeZone.secondsFromGMT())
        return addingTimeInterval(delta)
    }
}

extension UIColor {
    func rgb() -> Int? {
        var fRed : CGFloat = 0
        var fGreen : CGFloat = 0
        var fBlue : CGFloat = 0
        var fAlpha: CGFloat = 0
        if self.getRed(&fRed, green: &fGreen, blue: &fBlue, alpha: &fAlpha) {
            let iRed = Int(fRed * 255.0)
            let iGreen = Int(fGreen * 255.0)
            let iBlue = Int(fBlue * 255.0)
            let iAlpha = Int(fAlpha * 255.0)

            //  (Bits 24-31 are alpha, 16-23 are red, 8-15 are green, 0-7 are blue).
            let rgb = (iAlpha << 24) + (iRed << 16) + (iGreen << 8) + iBlue
            return rgb
        } else {
            // Could not extract RGBA components:
            return nil
        }
    }
    
    public convenience init?(hex: String) {
        let r, g, b, a: CGFloat

        if hex.hasPrefix("0x") {
            let start = hex.index(hex.startIndex, offsetBy: 2)
            let hexColor = String(hex[start...])

            if hexColor.count == 8 {
                let scanner = Scanner(string: hexColor)
                var hexNumber: UInt64 = 0

                if scanner.scanHexInt64(&hexNumber) {
                    a = CGFloat((hexNumber & 0xff000000) >> 24) / 255
                    r = CGFloat((hexNumber & 0x00ff0000) >> 16) / 255
                    g = CGFloat((hexNumber & 0x0000ff00) >> 8) / 255
                    b = CGFloat((hexNumber & 0x000000ff)) / 255
                    
                    self.init(red: r, green: g, blue: b, alpha: a)
                    return
                }
            }
        }

        return nil
    }

}
