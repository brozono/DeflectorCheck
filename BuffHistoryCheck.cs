namespace DeflectorCheck
{
    internal class BuffHistoryCheck
    {
        private static readonly DateTime UnixEpoch = new (1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        private static DateTime nowUTC;

        public static void SetNowUTC(DateTime now)
        {
            nowUTC = now;
        }

        public static DeflectorMetrics Check(Ei.ContractCoopStatusResponse.Types.ContributionInfo contributor, double startTime, double endTime)
        {
            DeflectorMetrics metrics = new ()
            {
                DeflectorSlotted = false,
                ActiveDeflectorSeconds = 0,
                InactiveDeflectorSeconds = 0,
                ActiveRatioSelf = 0,
                ActiveRatioCoop = 0,
                ActiveRatioContract = 0,
            };

            if (contributor.BuffHistory.Count > 0)
            {
                Ei.CoopBuffState lastBuffStatus = contributor.BuffHistory[^1];
                metrics.DeflectorSlotted = lastBuffStatus.HasEggLayingRate && lastBuffStatus.EggLayingRate > 1;

                bool deflectorActive = false;
                double startTimeActive = Math.Min(startTime, ServerTime(contributor.BuffHistory[0].ServerTimestamp));
                double startTimeInactive = Math.Min(startTime, ServerTime(contributor.BuffHistory[0].ServerTimestamp));

                foreach (Ei.CoopBuffState buffStatus in contributor.BuffHistory)
                {
                    if (buffStatus.HasEggLayingRate)
                    {
                        if (!deflectorActive && buffStatus.EggLayingRate > 1)
                        {
                            metrics.InactiveDeflectorSeconds += ServerTime(buffStatus.ServerTimestamp) - startTimeInactive;
                            startTimeActive = ServerTime(buffStatus.ServerTimestamp);
                            deflectorActive = true;
                        }
                        else if (deflectorActive && buffStatus.EggLayingRate <= 1)
                        {
                            metrics.ActiveDeflectorSeconds += ServerTime(buffStatus.ServerTimestamp) - startTimeActive;
                            startTimeInactive = ServerTime(buffStatus.ServerTimestamp);
                            deflectorActive = false;
                        }
                    }
                }

                if (deflectorActive && endTime > ServerTime(contributor.BuffHistory[^1].ServerTimestamp))
                {
                    metrics.ActiveDeflectorSeconds += endTime - ServerTime(contributor.BuffHistory[^1].ServerTimestamp);
                }

                metrics.ActiveRatioSelf = metrics.ActiveDeflectorSeconds / (metrics.ActiveDeflectorSeconds + metrics.InactiveDeflectorSeconds);
                if (double.IsNaN(metrics.ActiveRatioSelf))
                {
                    metrics.ActiveRatioSelf = 0;
                }
            }

            return metrics;
        }

        private static double FromDateTime(DateTime date) => (long)(date - UnixEpoch).TotalSeconds;

        private static double ServerTime(double seconds) => FromDateTime(nowUTC.AddSeconds(-1 * seconds));

        public struct DeflectorMetrics
        {
            public bool DeflectorSlotted;
            public double ActiveDeflectorSeconds;
            public double InactiveDeflectorSeconds;
            public double ActiveRatioSelf;
            public double ActiveRatioCoop;
            public double ActiveRatioContract;
        }
    }
}
