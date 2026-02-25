CREATE VIEW vwFact_ResourceAvailabilityPattern AS
SELECT
    ra.AvailabilityID,
    ra.ResourceID,
    ra.PatternType,
    ra.Mode,
    ra.AllDay,
    ra.StartTime,
    ra.EndTime,

    -- Pattern range (applies to weekly/monthly)
    pr.StartDate AS PatternStartDate,
    pr.EndDate AS PatternEndDate,
    pr.EndType As PatternEndType,
    pr.EndAfterOccurrences as PatternEndAfterOccurrences,

    -- Date range pattern
    pdr.StartDate AS RangeStartDate,
    pdr.EndDate AS RangeEndDate,

    -- Weekly pattern
    pw.RecurWeeks,
    pwd.DayOfWeek AS WeeklyDayOfWeek,

    -- Monthly pattern
    pm.MonthlyType,
    pm.DayOfMonth AS MonthlyDayOfMonth,
    pm.Ordinal AS MonthlyOrdinal,
    pm.DayOfWeek AS MonthlyDayOfWeek,
    pm.RecurMonths

FROM tblResourceAvailability ra
-- Pattern date range (1:1)
LEFT JOIN tblPatternRange pr
    ON ra.AvailabilityID = pr.AvailabilityID

-- Range pattern (1:1)
LEFT JOIN tblPatternDateRange pdr
    ON ra.AvailabilityID = pdr.AvailabilityID

-- Weekly pattern (1:1)
LEFT JOIN tblPatternWeekly pw
    ON ra.AvailabilityID = pw.AvailabilityID

-- Weekly days (1:many)
LEFT JOIN tblPatternWeeklyDays pwd
    ON ra.AvailabilityID = pwd.AvailabilityID

-- Monthly pattern (1:1)
LEFT JOIN tblPatternMonthly pm
    ON ra.AvailabilityID = pm.AvailabilityID;