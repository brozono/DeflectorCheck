namespace DeflectorCheck
{
    using EggIncApi;
    using Ei;

    internal class ContractDetail
    {
        private static Dictionary<string, ContractData> contracts = new ();
        private static string fileName;

        public static void LoadContractFile()
        {
            fileName = Config.GetCoopFolder() + Config.GroupName + "_ContractDetail.txt";

            if (Directory.Exists(Config.GetCoopFolder()))
            {
                if (File.Exists(fileName))
                {
                    try
                    {
                        string fileText = File.ReadAllText(fileName, System.Text.Encoding.UTF8);
                        contracts = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, ContractData>>(fileText);
                    }
                    catch
                    {
                        Logger.Warn("Failed to load Contract Detail from file");
                    }
                }
            }
        }

        public static void SaveContractFile()
        {
            if (!Directory.Exists(Config.GetCoopFolder()))
            {
                Directory.CreateDirectory(Config.GetCoopFolder());
            }

            try
            {
                var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(contracts, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(fileName, jsonString);
            }
            catch
            {
                Logger.Warn("Failed to save contract detail to file");
            }
        }

        public static bool AddContractCoop(string contract, string coop, string eggId)
        {
            EggIncFirstContactResponse firstContact;

            if (!HasCoop(contract, coop))
            {
                try
                {
                    var task = EggIncApi.GetFirstContact(eggId);
                    task.Wait();
                    firstContact = task.Result;

                    if (firstContact is null || firstContact.HasErrorCode || firstContact.HasErrorMessage)
                    {
                        firstContact = null;
                        return HasCoop(contract, coop);
                    }

                    FindDetail(firstContact.Backup.Contracts, contract, coop);
                    firstContact = null;
                }
                catch
                {
                    // TODO Log something...
                }
            }

            return HasCoop(contract, coop);
        }

        public static double GetCoopStartTime(string contract, string coop)
        {
            if (contracts.ContainsKey(contract))
            {
                if (contracts[contract].TimeAccepted.ContainsKey(coop))
                {
                    return contracts[contract].TimeAccepted[coop];
                }
            }

            return 0; // Hopefully we don't hit this
        }

        public static double GetExpectedDurationHours(string contract, double rate)
        {
            if (contracts.ContainsKey(contract))
            {
                double hours = contracts[contract].TargetEggs / rate;

                hours += contracts[contract].TokenTimerMinutes switch
                {
                    < 30 => 4,
                    < 120 => 6,
                    < 240 => 8,
                    _ => 10D,
                };
                return Config.TokenMultiplier() * hours;
            }

            // Default
            return Config.TokenMultiplier() * 24;
        }

        public static double GetCoopEndTime(string contract, string coop, double rate)
        {
            if (contracts.ContainsKey(contract))
            {
                if (contracts[contract].TimeAccepted.ContainsKey(coop))
                {
                    double hours = contracts[contract].TargetEggs / rate;

                    hours += contracts[contract].TokenTimerMinutes switch
                    {
                        < 30 => 4,
                        < 120 => 6,
                        < 240 => 8,
                        _ => 10D,
                    };

                    return contracts[contract].TimeAccepted[coop] + (hours * 3600);
                }
            }

            return 0; // Hopefully we don't hit this
        }

        public static double GetContractTarget(string contract)
        {
            if (contracts.ContainsKey(contract))
            {
                return contracts[contract].TargetEggs;
            }

            return 0;
        }

        public static uint GetContractSize(string contract)
        {
            if (contracts.ContainsKey(contract))
            {
                return contracts[contract].CoopSize;
            }

            return 0;
        }

        public static double GetContractTokenTimer(string contract)
        {
            if (contracts.ContainsKey(contract))
            {
                return contracts[contract].TokenTimerMinutes;
            }

            return 0;
        }

        public static List<string> GetSortedContracts()
        {
            var sortedContracts = from entry in contracts orderby entry.Value.TimeAccepted.First().Value descending select entry.Key;
            return sortedContracts.ToList();
        }

        public static bool HasCoop(string contract, string coop)
        {
            if (!contracts.ContainsKey(contract))
            {
                return false;
            }

            return contracts[contract].TimeAccepted.ContainsKey(coop);
        }

        private static void AddDetail(LocalContract localContract, string contractName, string coopName)
        {
            ContractData tempContractData;

            if (!contracts.ContainsKey(contractName))
            {
                tempContractData = new ()
                {
                    ContractId = contractName,
                    TargetEggs = localContract.Contract.Goals.Last().TargetAmount,
                    Egg = localContract.Contract.Egg,
                    CoopSize = localContract.Contract.MaxCoopSize,
                    TokenTimerMinutes = localContract.Contract.MinutesPerToken,
                    LengthSeconds = localContract.Contract.LengthSeconds,
                    ExpirationTime = localContract.Contract.ExpirationTime,
                    TimeAccepted = new (),
                };

                Logger.Info("Added Contract Detail: " + contractName);
                Logger.Info("Target Eggs: " + tempContractData.TargetEggs);
                Logger.Info("Eggs: " + tempContractData.Egg);
                Logger.Info("Coop Size: " + tempContractData.CoopSize);
                Logger.Info("Token Timer (min): " + tempContractData.TokenTimerMinutes);
                Logger.Info("Length: " + tempContractData.LengthSeconds);
                Logger.Info("Expiration: " + tempContractData.ExpirationTime);

                contracts.Add(contractName, tempContractData);
            }

            if (!contracts[contractName].TimeAccepted.ContainsKey(coopName))
            {
                contracts[contractName].TimeAccepted.Add(coopName, localContract.TimeAccepted);

                Logger.Info("Added Time Accepted Coop " + coopName + " to " + contractName + ": " + localContract.TimeAccepted);
            }
        }

        private static void FindDetail(MyContracts myContracts, string contractName, string coopName)
        {
            foreach (LocalContract localContract in myContracts.Contracts)
            {
                if (localContract.Contract.Identifier == contractName && localContract.CoopIdentifier == coopName)
                {
                    AddDetail(localContract, contractName, coopName);
                    return;
                }
            }

            foreach (LocalContract localContract in myContracts.Archive)
            {
                if (localContract.Contract.Identifier == contractName && localContract.CoopIdentifier == coopName)
                {
                    AddDetail(localContract, contractName, coopName);
                    return;
                }
            }
        }

        private struct ContractData
        {
            public string ContractId;
            public double TargetEggs;
            public Ei.Egg Egg;
            public uint CoopSize;
            public double TokenTimerMinutes;
            public double LengthSeconds;
            public double ExpirationTime;
            public Dictionary<string, double> TimeAccepted;
        }
    }
}
