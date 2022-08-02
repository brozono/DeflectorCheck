namespace DeflectorCheck
{
    internal class Config
    {
        private const string ConfigFileType = ".json";
        private const string RoamingFileName = "DeflectorCheckConfig";

        private static readonly string RoamingFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\DeflectorCheck\";
        private static readonly string RoamingFull = RoamingFolder + RoamingFileName + ConfigFileType;

        private static Dictionary<string, List<string>> coops;
        private static List<string> groupMembers;
        private static Dictionary<string, string> fuzzyMappings;
        private static string groupName;
        private static string eggIncID;
        private static Dictionary<string, List<string>> excludedCoops;
        private static bool slow;

        public static Dictionary<string, List<string>> Coops
        {
            get
            {
                return coops;
            }
        }

        public static List<string> GroupMembers
        {
            get { return groupMembers; }
        }

        public static string GroupName
        {
            get { return groupName; }
        }

        public static string EggIncID
        {
            get { return eggIncID; }
        }

        public static string GetConfigFileName()
        {
            return RoamingFull;
        }

        public static string GetFuzzyMapping(string value)
        {
            if (fuzzyMappings != null && fuzzyMappings.ContainsKey(value))
            {
                return fuzzyMappings[value];
            }

            return value;
        }

        public static bool IsExcluded(string contract, string coop)
        {
            if (excludedCoops != null &&
                excludedCoops.ContainsKey(contract) &&
                excludedCoops[contract].Contains(coop))
            {
                return true;
            }

            return false;
        }

        public static int TokenMultiplier()
        {
            if (slow)
            {
                return 2;
            }

            return 1;
        }

        public static string GetCoopFolder()
        {
            return RoamingFolder + GroupName + @"\";
        }

        public static void ParseConfig(string commandLineFileName)
        {
            coops = new Dictionary<string, List<string>>();
            groupMembers = new List<string>();
            string fileName = RoamingFull;

            if (commandLineFileName != string.Empty)
            {
                fileName = commandLineFileName;
            }

            if (!File.Exists(fileName))
            {
                Console.WriteLine("ERROR: Expected configuration file at: '" + fileName + "'");
                return;
            }

            try
            {
                string fileText = File.ReadAllText(fileName, System.Text.Encoding.UTF8);

                try
                {
                    var configData = Newtonsoft.Json.JsonConvert.DeserializeObject<ConfigStructure>(fileText);
                    coops = configData.Coops;
                    groupMembers = configData.GroupMembers;
                    fuzzyMappings = configData.FuzzyMappings;
                    groupName = configData.GroupName;
                    eggIncID = configData.EggIncID;
                    excludedCoops = configData.ExcludedCoops;
                    slow = configData.Slow;
                }
                catch (Exception e)
                {
                    Console.WriteLine("ERROR: Malformed config file");
                    Console.WriteLine(e.Message);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("ERROR: Error occured in parsing '" + RoamingFull + "'");
                Console.WriteLine(e.ToString());
            }
        }

        private class ConfigStructure
        {
            public Dictionary<string, List<string>> Coops { get; set; }

            public List<string> GroupMembers { get; set; }

            public Dictionary<string, string> FuzzyMappings { get; set; }

            public string GroupName { get; set; }

            public string EggIncID { get; set; }

            public Dictionary<string, List<string>> ExcludedCoops { get; set; }

            public bool Slow { get; set; }
        }
    }
}
