namespace Tempo.Exporter
{
    internal class TimeRange(TimeSpan from, TimeSpan to)
    {
        private TimeSpan From { get; set; } = from;
        private TimeSpan To { get; set; } = to;

        public TimeSpan Duration => To - From;

        internal static int CalculateBreaks(IEnumerable<TimeRange> timeRanges)
        {
            var sorted = timeRanges
                .OrderBy(tr => tr.From)
                .ToList();

            var totalBreak = TimeSpan.Zero;
            for (int i = 1; i < sorted.Count; i++)
            {
                var previous = sorted[i - 1];
                var current = sorted[i];
                var gap = current.From - previous.To;

                if (gap > TimeSpan.Zero)
                    totalBreak += gap;
            }

            return (int)totalBreak.TotalMinutes;
        }

        internal static TimeSpan GetWorkBegin(IEnumerable<TimeRange> timeRanges)
        {
            return timeRanges.Min(t => t.From);
        }

        internal static TimeSpan GetWorkEnd(IEnumerable<TimeRange> timeRanges)
        {
            return timeRanges.Max(t => t.To);   
        }
    }
}
