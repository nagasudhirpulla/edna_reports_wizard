using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VariableTime
{
    public class VarTime
    {
        public bool IsYearsVariable { get; set; } = false;
        public bool IsMonthsVariable { get; set; } = false;
        public bool IsDaysVariable { get; set; } = false;
        public bool IsHoursVariable { get; set; } = false;
        public bool IsMinutesVariable { get; set; } = false;
        public bool IsSecondsVariable { get; set; } = false;
        public int YearsOffset { get; set; } = 0;
        public int MonthsOffset { get; set; } = 0;
        public int DaysOffset { get; set; } = 0;
        public int HoursOffset { get; set; } = 0;
        public int MinutesOffset { get; set; } = 0;
        public int SecondsOffset { get; set; } = 0;

        public DateTime AbsoluteTime { get; set; } = DateTime.Now;

        internal VarTime Clone()
        {
            VarTime variableTime = new VarTime
            {
                IsYearsVariable = IsYearsVariable,
                IsMonthsVariable = IsMonthsVariable,
                IsDaysVariable = IsDaysVariable,
                IsHoursVariable = IsHoursVariable,
                IsMinutesVariable = IsMinutesVariable,
                IsSecondsVariable = IsSecondsVariable,
                YearsOffset = YearsOffset,
                MonthsOffset = MonthsOffset,
                DaysOffset = DaysOffset,
                HoursOffset = HoursOffset,
                MinutesOffset = MinutesOffset,
                SecondsOffset = SecondsOffset,
                AbsoluteTime = AbsoluteTime
            };
            return variableTime;
        }

        public DateTime GetTime()
        {
            DateTime absTime = AbsoluteTime;
            DateTime nowTime = DateTime.Now;

            // Make millisecond component as zero for the absolute time and now time
            absTime = absTime.AddMilliseconds(-1 * absTime.Millisecond);
            nowTime = nowTime.AddMilliseconds(-1 * nowTime.Millisecond);
            DateTime resultTime = nowTime;

            // Add offsets to current time as per the settings
            if (IsYearsVariable)
            {
                resultTime = resultTime.AddYears(YearsOffset);
            }
            if (IsMonthsVariable)
            {
                resultTime = resultTime.AddMonths(MonthsOffset);
            }
            if (IsDaysVariable)
            {
                resultTime = resultTime.AddDays(DaysOffset);
            }
            if (IsHoursVariable)
            {
                resultTime = resultTime.AddHours(HoursOffset);
            }
            if (IsMinutesVariable)
            {
                resultTime = resultTime.AddMinutes(MinutesOffset);
            }
            if (IsSecondsVariable)
            {
                resultTime = resultTime.AddSeconds(SecondsOffset);
            }

            // Set absolute time settings to the result time
            if (!IsYearsVariable)
            {
                resultTime = resultTime.AddYears(absTime.Year - resultTime.Year);
            }
            if (!IsMonthsVariable)
            {
                resultTime = resultTime.AddMonths(absTime.Month - resultTime.Month);
            }
            if (!IsDaysVariable)
            {
                resultTime = resultTime.AddDays(absTime.Day - resultTime.Day);
            }
            if (!IsHoursVariable)
            {
                resultTime = resultTime.AddHours(absTime.Hour - resultTime.Hour);
            }
            if (!IsMinutesVariable)
            {
                resultTime = resultTime.AddMinutes(absTime.Minute - resultTime.Minute);
            }
            if (!IsSecondsVariable)
            {
                resultTime = resultTime.AddSeconds(absTime.Second - resultTime.Second);
            }
            return resultTime;
        }
    }
}
