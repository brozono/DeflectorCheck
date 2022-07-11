namespace DeflectorCheck
{
    internal class Logger
    {
        private static readonly NLog.Logger LoggerValue = NLog.LogManager.GetCurrentClassLogger();

        public static void ConfigureLogger()
        {
            NLog.LogManager.ThrowExceptions = true;

            NLog.Config.LoggingConfiguration config = new ();

            NLog.Layouts.JsonLayout layout = new ()
            {
                Attributes =
                {
                    new NLog.Layouts.JsonAttribute("time", "${longdate}", false),
                    new NLog.Layouts.JsonAttribute("level", "${level:upperCase=true}", false),
                    new NLog.Layouts.JsonAttribute("message", "${message}", false),
                    new NLog.Layouts.JsonAttribute("exception", "${exception:format=ToString}", false),
                },
            };

            NLog.Targets.FileTarget fileTarget = new ("logfile")
            {
                FileName = "${specialfolder:folder=ApplicationData}/DeflectorCheck/DeflectorCheckLog.txt",
                Layout = layout,
                KeepFileOpen = true,
                OpenFileCacheTimeout = 30,
                AutoFlush = false,
                OpenFileFlushTimeout = 10,
                ArchiveAboveSize = 10240000,
                MaxArchiveFiles = 10,
            };

            NLog.Common.InternalLogger.LogToConsole = false;
            NLog.Common.InternalLogger.LogLevel = NLog.LogLevel.Trace;

            config.AddRule(NLog.LogLevel.Info, NLog.LogLevel.Fatal, fileTarget);

            NLog.LogManager.Configuration = config;

            LoggerValue.Info("Logging Started");
        }

        public static void SetLogLevel(string toLevel)
        {
            NLog.LogLevel logLevelToSet;

            try
            {
                logLevelToSet = NLog.LogLevel.FromString(toLevel);
            }
            catch (ArgumentException)
            {
                logLevelToSet = NLog.LogLevel.Info;
            }

            foreach (NLog.Config.LoggingRule rule in NLog.LogManager.Configuration.LoggingRules)
            {
                rule.SetLoggingLevels(logLevelToSet, NLog.LogLevel.Fatal);
            }

            NLog.LogManager.ReconfigExistingLoggers();
        }

        public static string GetLogLevel()
        {
            NLog.LogLevel level = NLog.LogLevel.Fatal;

            foreach (NLog.Config.LoggingRule rule in NLog.LogManager.Configuration.LoggingRules)
            {
                foreach (NLog.LogLevel ruleLevel in rule.Levels)
                {
                    if (ruleLevel < level)
                    {
                        level = ruleLevel;
                    }
                }
            }

            return level.ToString();
        }

        public static void Log(NLog.LogLevel level, string message, params object[] args)
        {
            LoggerValue.Log(level, message, args);
        }

        public static void Log(NLog.LogLevel level, Exception exception, string message, params object[] args)
        {
            LoggerValue.Log(level, exception, message, args);
        }

        public static void Trace(string message, params object[] args)
        {
            LoggerValue.Trace(message, args);
        }

        public static void Trace(string message, Exception exception, params object[] args)
        {
            LoggerValue.Trace(message, exception, args);
        }

        public static void Debug(string message, params object[] args)
        {
            LoggerValue.Debug(message, args);
        }

        public static void Debug(string message, Exception exception, params object[] args)
        {
            LoggerValue.Debug(message, exception, args);
        }

        public static void Info(string message, params object[] args)
        {
            LoggerValue.Info(message, args);
        }

        public static void Info(string message, Exception exception, params object[] args)
        {
            LoggerValue.Info(message, exception, args);
        }

        public static void Warn(string message, params object[] args)
        {
            LoggerValue.Warn(message, args);
        }

        public static void Warn(string message, Exception exception, params object[] args)
        {
            LoggerValue.Warn(message, exception, args);
        }

        public static void Error(string message, params object[] args)
        {
            LoggerValue.Warn(message, args);
        }

        public static void Error(string message, Exception exception, params object[] args)
        {
            LoggerValue.Warn(message, exception, args);
        }

        public static void Fatal(string message, params object[] args)
        {
            LoggerValue.Fatal(message, args);
        }

        public static void Fatal(string message, Exception exception, params object[] args)
        {
            LoggerValue.Fatal(message, exception, args);
        }
    }
}
