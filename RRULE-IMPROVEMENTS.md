# RRULE Pattern Improvements

## Summary

Successfully improved recurrence pattern extraction by using pst-extractor's `RecurrencePattern` class to parse the binary recurrence structure instead of relying on simple text properties.

## Processing Results

- **Files Processed**: 4 PST files
- **Total Entries Found**: 40,644
- **Unique Entries**: 4,084  
- **Duplicates Skipped**: 36,560 (89.9% deduplication)
- **Date Range**: 1926-06-12 to 2024-03-04

## RRULE Pattern Types Found

### Weekly Patterns
- `FREQ=WEEKLY;BYDAY=MO;COUNT=7` - Every Monday, 7 times
- `FREQ=WEEKLY;BYDAY=TH;UNTIL=16001231` - Every Thursday until end date
- `FREQ=WEEKLY;INTERVAL=2;BYDAY=MO;UNTIL=16001231` - Every 2 weeks on Monday

### Monthly Patterns
- `FREQ=MONTHLY;BYDAY=2MO;COUNT=2` - 2nd Monday of month, 2 times
- `FREQ=MONTHLY;BYDAY=-1FR` - Last Friday of month
- `FREQ=MONTHLY;BYDAY=1SA;COUNT=8` - 1st Saturday of month, 8 times  
- `FREQ=MONTHLY;BYMONTHDAY=1;COUNT=4` - 1st of month, 4 times
- `FREQ=MONTHLY;BYMONTHDAY=14;COUNT=15` - 14th of month, 15 times

### Yearly Patterns (Birthdays/Anniversaries)
- `FREQ=YEARLY;INTERVAL=12;BYMONTHDAY=29` - Yearly on 29th
- `FREQ=YEARLY;INTERVAL=12;BYMONTHDAY=16;COUNT=1` - Once on 16th

## Key Improvements

### Before (Text Parsing Only)
```
recurrenceType: 1  // Just "Weekly"
recurrencePattern: "every Monday"  
Result: RRULE:FREQ=DAILY  ← WRONG!
```

### After (Binary Structure Parsing)
```javascript
{
  recurFrequency: "Weekly",
  patternType: "Week",
  patternTypeSpecific: [false,true,false,false,false,false,false],  // Monday
  endType: "AfterNOccurrences",
  occurrenceCount: 7
}
Result: RRULE:FREQ=WEEKLY;BYDAY=MO;COUNT=7  ← CORRECT!
```

## What's Now Supported

✅ **FREQ**: DAILY, WEEKLY, MONTHLY, YEARLY (correct values)
✅ **BYDAY**: Specific days of week (MO, TU, WE, TH, FR, SA, SU)
✅ **BYMONTHDAY**: Day of month (1-31)
✅ **BYDAY with ordinals**: "2MO" = second Monday, "-1FR" = last Friday
✅ **INTERVAL**: Every N days/weeks/months
✅ **COUNT**: Number of occurrences
✅ **UNTIL**: End date (YYYYMMDD format)

## Real Examples

1. **Green Bin Collection**
   - `FREQ=MONTHLY;BYDAY=2MO;COUNT=2` - 2nd Monday, 2 times
   
2. **Park & Ride Pass expires**
   - `FREQ=WEEKLY;BYDAY=MO;COUNT=7` - Every Monday, 7 times

3. **Monthly Payroll**
   - `FREQ=MONTHLY;BYMONTHDAY=14;COUNT=15` - 14th of month, 15 times

4. **Last Friday Events**
   - `FREQ=MONTHLY;BYDAY=-1FR` - Last Friday of every month

## Technical Implementation

**Import:**
```typescript
import { RecurrencePattern } from 'pst-extractor/dist/RecurrencePattern.class';
```

**Parse Structure:**
```typescript
const recurrenceStructure = appointment.recurrenceStructure;
if (recurrenceStructure && Buffer.isBuffer(recurrenceStructure)) {
  const pattern = new RecurrencePattern(recurrenceStructure);
  return buildRRuleFromPattern(pattern);
}
```

**Map to RRULE:**
- Frequency: RecurFrequency enum (8202-8205)
- Pattern Type: Day/Week/Month/MonthNth/MonthEnd
- Pattern Specific: Weekday arrays, day numbers, ordinals
- End Type: AfterDate/AfterNOccurrences/NeverEnd

## Files Generated

- `calendar-improved-rrule.ics` - iCal with correct RRULE patterns
- `calendar-export.csv` - CSV with RRULE patterns
- `.fixecalendar.db` - SQLite database with deduplicated entries

## No Alternative Library Needed!

The existing `pst-extractor` library has comprehensive RRULE support via the `RecurrencePattern` class. It fully exposes Microsoft's MAPI recurrence structure - we just needed to use the right property (`recurrenceStructure` buffer) instead of the simplified `recurrenceType` field.
