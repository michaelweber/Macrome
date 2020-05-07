using System;
using System.Text;
using System.Diagnostics;
using b2xtranslator.Tools;
using System.Reflection;
using System.IO;

namespace b2xtranslator.Shell
{
    public class CommandLineTranslator
    {
        // parsed arguments
        public static string InputFile;
        public static string ChoosenOutputFile;

        public static void InitializeLogger()
        {
            // let the Console listen to the Trace messages
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
        }

        /// <summary>
        /// Prints the heading row of the tool
        /// </summary>
        public static void PrintWelcome(string toolname, string revisionResource)
        {
            bool backup = TraceLogger.EnableTimeStamp;
            TraceLogger.EnableTimeStamp = false;
            var welcome = new StringBuilder();
            welcome.Append("Welcome to ");
            welcome.Append(toolname);
            welcome.Append(" (r");
            welcome.Append(getRevision(revisionResource));
            welcome.Append(")\n");
            TraceLogger.Simple(welcome.ToString());
            TraceLogger.EnableTimeStamp = backup;
        }

        /// <summary>
        /// Prints the usage row of the tool
        /// </summary>
        public static void PrintUsage(string toolname)
        {
            var usage = new StringBuilder();
            usage.AppendLine("Usage: " + toolname + " [-c | inputfile] [-o outputfile] [-v level] [-?]");
            usage.AppendLine("-o <outputfile>  change output filename");
            usage.AppendLine("-v <level>     set trace level, where <level> is one of the following:");
            usage.AppendLine("               none (0)    print nothing");
            usage.AppendLine("               error (1)   print all errors");
            usage.AppendLine("               warning (2) print all errors and warnings");
            usage.AppendLine("               info (3)    print all errors, warnings and infos (default)");
            usage.AppendLine("               debug (4)   print all errors, warnings, infos and debug messages");
            usage.AppendLine("-?             print this help");
            Console.WriteLine(usage.ToString());
        }

        /// <summary>
        /// Returns the revision that is stored in the embedded resource "revision.txt".
        /// Returns -1 if something goes wrong
        /// </summary>
        /// <returns></returns>
        private static int getRevision(string revisionResource)
        {
            int rev = -1;

            try
            {
                var a = Assembly.GetEntryAssembly();
                using (var s = a.GetManifestResourceStream(revisionResource))
                using (var reader = new StreamReader(s))
                {
                    rev = int.Parse(reader.ReadLine());
                    s.Close();
                }
            }
            catch (Exception) { }

            return rev;
        }

        //public static RegistryKey GetContextMenuKey(string triggerExtension, string contextMenuText)
        //{
        //    RegistryKey result = null;
        //    try
        //    {
        //        string defaultWord = (string)Registry.ClassesRoot.OpenSubKey(triggerExtension).GetValue("");
        //        result = Registry.ClassesRoot.CreateSubKey(defaultWord).CreateSubKey("shell").CreateSubKey(contextMenuText);
        //    }
        //    catch (Exception)
        //    {
        //    }
        //    return result;
        //}

        //public static void RegisterForContextMenu(RegistryKey entryKey)
        //{
        //    if (entryKey != null)
        //    {
        //        RegistryKey convertCommand = entryKey.CreateSubKey("Command");
        //        convertCommand.SetValue("", String.Format("\"{0}\" \"%1\"", Assembly.GetCallingAssembly().Location));
        //    }
        //}

        /// <summary>
        /// Parses the arguments of the tool
        /// </summary>
        /// <param name="args">The args array</param>
        public static void ParseArgs(string[] args, string toolName)
        {
            try
            {
                if (args[0] == "-?")
                {
                    PrintUsage(toolName);
                    Environment.Exit(0);
                }
                else
                {
                    InputFile = args[0];
                }

                for (int i = 1; i < args.Length; i++)
                {
                    if (args[i].ToLower() == "-v")
                    {
                        //parse verbose level
                        string verbose = args[i + 1].ToLower();
                        int vLvl;
                        if (int.TryParse(verbose, out vLvl))
                        {
                            TraceLogger.LogLevel = (TraceLogger.LoggingLevel)vLvl;
                        }
                        else if (verbose == "error")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Error;
                        }
                        else if (verbose == "warning")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Warning;
                        }
                        else if (verbose == "info")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Info;
                        }
                        else if (verbose == "debug")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.Debug;
                        }
                        else if (verbose == "none")
                        {
                            TraceLogger.LogLevel = TraceLogger.LoggingLevel.None;
                        }
                    }
                    else if (args[i].ToLower() == "-o")
                    {
                        //parse output file name
                        ChoosenOutputFile = args[i + 1];
                    }
                }
            }
            catch (Exception)
            {
                TraceLogger.Error("At least one of the required arguments was not correctly set.\n");
                PrintUsage(toolName);
                Environment.Exit(1);
            }
        }
    }
}
