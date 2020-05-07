using System;

namespace b2xtranslator.Tools
{
    public static class TraceLogger
    {
        public static bool EnableTimeStamp = true;

        public enum LoggingLevel
        {
            None = 0,
            Error,
            Warning,
            Info,
            Debug,
            DebugInternal
        }

        static LoggingLevel _logLevel = LoggingLevel.Info;
        public static LoggingLevel LogLevel
        {
            get { return TraceLogger._logLevel; }
            set { TraceLogger._logLevel = value; }
        }


        private static void WriteLine(string msg, LoggingLevel level)
        {
            if (_logLevel >= level && EnableTimeStamp)
            {
                try
                {
                    System.Diagnostics.Trace.WriteLine(string.Format("{0} " + msg, System.DateTime.Now));
                }
                catch (Exception)
                {
                    System.Diagnostics.Trace.WriteLine("The tracing of the following message threw an error: " + msg);
                }
            }
            else if (_logLevel >= level)
            {
                System.Diagnostics.Trace.WriteLine(msg);
            }
        }

        /// <summary>
        /// Write a line on error level (is written if level != none)
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="objs"></param>
        public static void Simple(string msg, params object[] objs)
        {
            if (msg == null || msg == "")
                return;

            WriteLine(string.Format(msg, objs), LoggingLevel.Error);
        }

        public static void DebugInternal(string msg, params object[] objs)
        {
            if (msg == null || msg == "")
                return;

            WriteLine("[D] " + string.Format(msg, objs), LoggingLevel.DebugInternal);
        }

        public static void Debug(string msg, params object[] objs)
        {
            if (msg == null || msg == "")
                return;

            WriteLine("[D] " + string.Format(msg, objs), LoggingLevel.Debug);
        }


        public static void Info(string msg, params object[] objs)
        {
            if (msg == null || msg == "")
                return;

            WriteLine("[I] " + string.Format(msg, objs), LoggingLevel.Info);
        }


        public static void Warning(string msg, params object[] objs)
        {
            if (msg == null || msg == "")
                return;

            WriteLine("[W] " + string.Format(msg, objs), LoggingLevel.Warning);
        }


        public static void Error(string msg, params object[] objs)
        {
            if (msg == null || msg == "")
                return;

            WriteLine("[E] " + string.Format(msg, objs), LoggingLevel.Error);
        }
    }
}

