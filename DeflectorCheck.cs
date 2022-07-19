﻿using Ei;

namespace DeflectorCheck
{
#pragma warning disable CS1591 // Missing XML comment for publicly visible type or member
    public class DeflectorCheck
    {
        private static bool debug = false;
        private static Dictionary<string, List<string>> coops;
        private static int indent;

        public static void Main(string[] arguments)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Logger.ConfigureLogger();

            string commandLineFileName = string.Empty;
            if (arguments.Length > 0)
            {
                commandLineFileName = arguments[0];
            }

            Config.ParseConfig(commandLineFileName);
            coops = Config.Coops;
            Debug();

            indent = 0;

            ContractCoopStatusResponse coopStatus;

            GroupData.AddMembers();

            foreach (string contract in coops.Keys)
            {
                Logger.Info("Contract: " + contract);
                Console.WriteLine(new string(' ', indent) + "Contract: " + contract);

                double totalActiveDeflectorSecondsContract = 0;
                int contributorCountContract = 0;

                indent += 2;
                foreach (string coop in coops[contract])
                {
                    Logger.Info("Coop: " + coop);
                    Console.WriteLine(new string(' ', indent) + "Coop: " + coop);
                    indent += 2;

                    try
                    {
                        coopStatus = GetCoop(contract, coop, Config.EggIncID);
                    }
                    catch
                    {
                        Logger.Error("Could not get Coop Status for " + contract + "::" + coop);
                        indent -= 2;
                        continue;
                    }

                    if (!coopStatus.HasCreatorId)
                    {
                        Logger.Error("Could not get valid Coop Status for " + contract + "::" + coop);
                        indent -= 2;
                        continue;
                    }

                    double totalEggsPerHour = 0;

                    ContractDetail.AddContractCoop(contract, coop, coopStatus.CreatorId);

                    foreach (var contributor in coopStatus.Contributors)
                    {
                        totalEggsPerHour += contributor.ContributionRate * 3600;
                    }

                    double totalActiveDeflectorSecondsCoop = 0;
                    int contributorCountCoop = 0;
                    string logIdentifier = contract + "::" + coop;

                    Logger.Debug(logIdentifier + " Start Time: " + ContractDetail.GetCoopStartTime(contract, coop));
                    Logger.Debug(logIdentifier + " End Time: " + ContractDetail.GetCoopEndTime(contract, coop, totalEggsPerHour));

                    foreach (var contributor in coopStatus.Contributors)
                    {
                        Logger.Info(logIdentifier + " Contributor: " + contributor.UserName);
                        indent += 2;

                        string memberMatch = GroupData.MatchMember(contributor.UserName);

                        if (memberMatch != string.Empty)
                        {
                            Logger.Info(logIdentifier + " Matched User Name '" + contributor.UserName + "' to '" + memberMatch + "'");

                            string logIdentifierDetail = logIdentifier + "::" + contributor.UserName;

                            BuffHistoryCheck.DeflectorMetrics metrics = BuffHistoryCheck.Check(
                                contributor,
                                ContractDetail.GetCoopStartTime(contract, coop),
                                ContractDetail.GetCoopEndTime(contract, coop, totalEggsPerHour));
                            GroupData.SetMemberData(logIdentifierDetail, memberMatch, contributor.UserName, contract, metrics);
                            totalActiveDeflectorSecondsCoop += metrics.ActiveDeflectorSeconds;
                            contributorCountCoop++;
                        }
                        else
                        {
                            Logger.Warn(logIdentifier + " No Match for User Name '" + contributor.UserName + "'");
                        }

                        indent -= 2;
                    }

                    double averageActiveDeflectorSeconds = totalActiveDeflectorSecondsCoop / contributorCountCoop;

                    foreach (var contributor in coopStatus.Contributors)
                    {
                        GroupData.CalculateCoopRatio(contributor.UserName, contract, averageActiveDeflectorSeconds);
                    }

                    totalActiveDeflectorSecondsContract += totalActiveDeflectorSecondsCoop;
                    contributorCountContract += contributorCountCoop;

                    indent -= 2;
                }

                double averageActiveDeflectorSecondsContract = totalActiveDeflectorSecondsContract / contributorCountContract;
                GroupData.CalculateContractRatio(contract, averageActiveDeflectorSecondsContract);

                indent -= 2;
            }

            GroupData.CheckData();
            GroupData.Log();

            Spreadsheet spreadsheet = new ();
            spreadsheet.Create();

            Console.WriteLine("Press any key when ready to exit...");
            Console.ReadKey();
        }

        private static ContractCoopStatusResponse GetCoop(string contract, string coop, string id)
        {
            var task = EggIncApi.EggIncApi.GetCoopStatus(contract, coop, id);
            task.Wait();
            return task.Result;
        }

        private static void Debug()
        {
            if (debug)
            {
                List<string> keys = coops.Keys.ToList<string>();
                foreach (string key in keys)
                {
                    if (key != "preparations-2022")
                    {
                        coops.Remove(key);
                    }
                }
            }
        }
    }
#pragma warning restore CS1591 // Missing XML comment for publicly visible type or member
}