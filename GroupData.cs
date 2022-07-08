namespace DeflectorCheck
{
    internal class GroupData
    {
        private static readonly Dictionary<string, MemberData> DataValue = new ();

        public static void AddMembers()
        {
            List<string> groupMembers = Config.GroupMembers;

            foreach (string member in groupMembers)
            {
                DataValue.Add(member, new MemberData() { EiUserName = string.Empty, CoopMetrics = new () });
            }
        }

        public static string MatchMember(string match)
        {
            string fuzzyMatch = Config.GetFuzzyMapping(match);
            string fuzzyMember;

            foreach (string member in DataValue.Keys)
            {
                fuzzyMember = Config.GetFuzzyMapping(member);
                if (fuzzyMatch.Contains(fuzzyMember))
                {
                    return member;
                }
            }

            return string.Empty;
        }

        public static void SetMemberData(string logPreamble, string member, string userName, string contract, BuffHistoryCheck.DeflectorMetrics metrics)
        {
            if (metrics.DeflectorSlotted)
            {
                Logger.Info(logPreamble + " Deflecting");
            }
            else
            {
                Logger.Info(logPreamble + " NOT Deflecting");
            }

            Logger.Info(logPreamble + " Deflector Active for " + ReadableElapsedTime(metrics.ActiveDeflectorSeconds));
            Logger.Info(logPreamble + " Deflector Inactive for " + ReadableElapsedTime(metrics.InactiveDeflectorSeconds));
            Logger.Info(logPreamble + " Deflector Ratio " + metrics.ActiveRatioSelf.ToString("P"));

            MemberData temp = DataValue[member];

            temp.EiUserName = userName;
            temp.CoopMetrics.Add(contract, metrics);
            DataValue[member] = temp;
        }

        public static void CalculateCoopRatio(string match, string contract, double coopAverage)
        {
            string member = MatchMember(match);

            if (member != string.Empty)
            {
                if (DataValue[member].CoopMetrics.ContainsKey(contract))
                {
                    BuffHistoryCheck.DeflectorMetrics tempMetrics = DataValue[member].CoopMetrics[contract];
                    tempMetrics.ActiveRatioCoop = tempMetrics.ActiveDeflectorSeconds / coopAverage;

                    if (double.IsNaN(tempMetrics.ActiveRatioCoop))
                    {
                        tempMetrics.ActiveRatioCoop = 0;
                    }

                    DataValue[member].CoopMetrics[contract] = tempMetrics;
                }
            }
        }

        public static void CalculateContractRatio(string contract, double contractAverage)
        {
            foreach (string member in DataValue.Keys)
            {
                if (DataValue[member].CoopMetrics.ContainsKey(contract))
                {
                    BuffHistoryCheck.DeflectorMetrics tempMetrics = DataValue[member].CoopMetrics[contract];
                    tempMetrics.ActiveRatioContract = tempMetrics.ActiveDeflectorSeconds / contractAverage;

                    if (double.IsNaN(tempMetrics.ActiveRatioContract))
                    {
                        tempMetrics.ActiveRatioContract = 0;
                    }

                    DataValue[member].CoopMetrics[contract] = tempMetrics;
                }
            }
        }

        public static void CheckData()
        {
            Console.WriteLine("Checking Group Data...");
            Logger.Info("Checking Group Data...");

            foreach (var item in DataValue)
            {
                var value = item.Value;
                value.EiUserName = Config.GetFuzzyMapping(value.EiUserName);

                if (value.EiUserName == string.Empty)
                {
                    Logger.Warn("Did not find '" + item.Key + "' in any coop!");
                    Console.WriteLine("  Did not find '" + item.Key + "' in any coop!");
                    value.EiUserName = Config.GetFuzzyMapping(item.Key);
                }

                DataValue[item.Key] = value;
            }
        }

        public static void Log()
        {
            List<string> sortedData = (from entry in DataValue orderby entry.Value.EiUserName ascending select entry.Key).ToList<string>();

            foreach (string member in sortedData)
            {
                Logger.Info("Group Member: " + member);
                Logger.Info("Egg Name: " + DataValue[member].EiUserName);
                foreach (var item in DataValue[member].CoopMetrics)
                {
                    Logger.Info("Coop: " + item.Key);
                    Logger.Info(item.Key + " Deflecting: " + item.Value.DeflectorSlotted);
                    Logger.Info(item.Key + " Self Ratio: " + item.Value.ActiveRatioSelf.ToString("P"));
                    Logger.Info(item.Key + " Coop Ratio: " + item.Value.ActiveRatioCoop.ToString("P"));
                    Logger.Info(item.Key + " Coop Ratio: " + item.Value.ActiveRatioContract.ToString("P"));
                }
            }
        }

        public static Dictionary<string, MemberData> Data()
        {
            return DataValue;
        }

        public static List<string> GetSortedMembers()
        {
            var sortedMembers = from entry in DataValue orderby entry.Value.EiUserName ascending select entry.Key;
            return sortedMembers.ToList();
        }

        public static string MemberEggName(string member)
        {
            if (DataValue.ContainsKey(member))
            {
                return DataValue[member].EiUserName;
            }

            return string.Empty;
        }

        public static bool MemberHasContract(string member, string contract)
        {
            if (DataValue.ContainsKey(member))
            {
                return DataValue[member].CoopMetrics.ContainsKey(contract);
            }

            return false;
        }

        public static bool MemberPassesQualifiers(string member, string contract, double personalThreshold, double coopThreshold, double contractThreshold)
        {
            if (MemberHasContract(member, contract))
            {
                return DataValue[member].CoopMetrics[contract].DeflectorSlotted ||
                       DataValue[member].CoopMetrics[contract].ActiveRatioSelf > personalThreshold ||
                       DataValue[member].CoopMetrics[contract].ActiveRatioCoop > coopThreshold ||
                       DataValue[member].CoopMetrics[contract].ActiveRatioContract > contractThreshold;
            }

            return false;
        }

        public static bool MemberDeflecting(string member, string contract)
        {
            if (MemberHasContract(member, contract))
            {
                return DataValue[member].CoopMetrics[contract].DeflectorSlotted;
            }

            return false;
        }

        public static double MemberPersonalRatio(string member, string contract)
        {
            if (MemberHasContract(member, contract))
            {
                return DataValue[member].CoopMetrics[contract].ActiveRatioSelf;
            }

            return 0;
        }

        public static double MemberCoopRatio(string member, string contract)
        {
            if (MemberHasContract(member, contract))
            {
                return DataValue[member].CoopMetrics[contract].ActiveRatioCoop;
            }

            return 0;
        }

        public static double MemberContractRatio(string member, string contract)
        {
            if (MemberHasContract(member, contract))
            {
                return DataValue[member].CoopMetrics[contract].ActiveRatioContract;
            }

            return 0;
        }

        private static string ReadableElapsedTime(double elapsed)
        {
            int days = (int)(elapsed / 86400);
            elapsed %= 86400;
            int hours = (int)(elapsed / 3600);
            elapsed %= 3600;
            int minutes = (int)(elapsed / 60);
            int seconds = (int)(elapsed % 60);

            return days.ToString() + "::" + hours.ToString() + "::" + minutes.ToString() + "::" + seconds.ToString();
        }

        public struct MemberData
        {
            public string EiUserName;
            public Dictionary<string, BuffHistoryCheck.DeflectorMetrics> CoopMetrics;
        }
    }
}
