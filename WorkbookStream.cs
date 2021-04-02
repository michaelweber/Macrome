using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Common;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.xls.XlsFileFormat;
using b2xtranslator.xls.XlsFileFormat.Records;

namespace Macrome
{
    public class WorkbookStream
    {
        private List<BiffRecord> _biffRecords;

        public List<BiffRecord> Records
        {
            get { return _biffRecords; }
        }

        public WorkbookStream(string filePath)
        {
            using (var fs = new FileStream(filePath, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                try
                {
                    var wbStream = ssr.GetStream("Workbook");
                    byte[] wbBytes = new byte[wbStream.Length];
                    wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                    _biffRecords = RecordHelper.ParseBiffStreamBytes(wbBytes);
                }
                catch (StreamNotFoundException)
                {
                    var wbStream = ssr.GetStream("Book");
                    Console.WriteLine("WARNING: Main stream is in a Book record indicating legacy Excel 5 BIFF format. This may not parse correctly.");

                    byte[] wbBytes = new byte[wbStream.Length];
                    wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                    try
                    {
                        _biffRecords = RecordHelper.ParseBiffStreamBytes(wbBytes);
                    }
                    catch (Exception)
                    {
                        throw new NotImplementedException("Error parsing Book stream: Macrome currently doesn't support the Excel 5 BIFF format.");
                    }
                }
            }
        }

        public WorkbookStream(List<BiffRecord> records)
        {
            _biffRecords = records;
        }

        public WorkbookStream(byte[] workbookBytes)
        {
            _biffRecords = RecordHelper.ParseBiffStreamBytes(workbookBytes);
        }

        public WorkbookStream RemoveRecord(BiffRecord recordToRemove)
        {
            if (ContainsRecord(recordToRemove) == false)
            {
                throw new ArgumentException("Could not find recordToRemove");
            }

            var removeRecordOffset = GetRecordOffset(recordToRemove);

            var newRecords = _biffRecords.Take(removeRecordOffset).Concat(
                _biffRecords.TakeLast(_biffRecords.Count - removeRecordOffset - 1)).ToList();

            return new WorkbookStream(newRecords);
        }

        public int GetFirstEmptyRowInColumn(int col)
        {
            List<Formula> colFormulas = GetAllRecordsByType<Formula>().Where(f => f.col == col).ToList();
            int maxRwVal = 0;
            foreach (var colFormula in colFormulas)
            {
                if (colFormula.rw > maxRwVal) maxRwVal = colFormula.rw;
            }

            return maxRwVal;
        }

        public WorkbookStream InsertRecord(BiffRecord recordToInsert, BiffRecord insertAfterRecord = null)
        {
            return InsertRecords(new List<BiffRecord>() {recordToInsert}, insertAfterRecord);
        }

        public WorkbookStream InsertRecords(List<BiffRecord> recordsToInsert, BiffRecord insertAfterRecord = null)
        {
            if (insertAfterRecord == null)
            {
                List<BiffRecord> recordsWithAppendedRecord = _biffRecords.Concat(recordsToInsert).ToList();
                return new WorkbookStream(recordsWithAppendedRecord);
            }

            if (ContainsRecord(insertAfterRecord) == false)
            {
                throw new ArgumentException("Could not find insertAfterRecord");
            }

            var insertRecordOffset = GetRecordOffset(insertAfterRecord) + 1;

            //records [r1, TARGET, r2, r3, r4, r5]
            //records.count = 6
            //insertRecordOffset = 2
            //records.Take(2) = [r1, TARGET]
            //records.TakeLast(4) = [r2, r3, r4, r5]
            //output = [r1, TARGET, INSERT, r2, r3, r4, r5]

            var newRecords = _biffRecords.Take(insertRecordOffset).Concat(recordsToInsert)
                .Concat(_biffRecords.TakeLast(_biffRecords.Count - insertRecordOffset)).ToList();

            return new WorkbookStream(newRecords);
        }

        public bool ContainsRecord(BiffRecord record)
        {
            var matchingRecordTypes = _biffRecords.Where(r => r.Id == record.Id).ToList();
            return matchingRecordTypes.Any(r => r.Equals(record));
        }


        public List<BiffRecord> GetRecordsForBOFRecord(BOF sheetBeginRecord)
        {
            var sheetRecords = _biffRecords.SkipWhile(r => r.Equals(sheetBeginRecord) == false).ToList();

            int sheetSize = sheetRecords.TakeWhile(r => r.Id != RecordType.EOF).Count() + 1;

            return sheetRecords.Take(sheetSize).ToList();
        }

        public WorkbookStream ReplaceRecord(BiffRecord oldRecord, BiffRecord newRecord)
        {
            if (ContainsRecord(oldRecord) == false)
            {
                throw new ArgumentException("Could not find oldRecord");
            }

            //records [r1, OLD, r2, r3, r4, r5]
            //records.count = 6
            //replaceRecordOffset = 1
            //records.Take(1) = [r1]
            //records.TakeLast(4) = [r2, r3, r4, r5]
            //output = [r1, NEW, r2, r3, r4, r5]

            var replaceRecordOffset = GetRecordOffset(oldRecord);

            var newRecords = _biffRecords.Take(replaceRecordOffset).Append(newRecord)
                .Concat(_biffRecords.TakeLast(_biffRecords.Count - (replaceRecordOffset + 1))).ToList();

            return new WorkbookStream(newRecords);
        }

        public int GetLabelOffset(string labelName)
        {
            List<Lbl> labels = GetAllRecordsByType<Lbl>();

            int offset = 1;
            foreach (var label in labels)
            {
                if (label.Name.Equals(labelName)) return offset;

                offset += 1;
            }

            throw new ArgumentException(string.Format("Cannot find Lbl record with name {0}", labelName));
        }

        public WorkbookStream AddSheet(BoundSheet8 sheetHeader, byte[] sheetBytes)
        {
            WorkbookStream newStream = new WorkbookStream(Records);
            List<BoundSheet8> existingBoundSheets = newStream.GetAllRecordsByType<BoundSheet8>();
            BoundSheet8 lastSheet8 = existingBoundSheets.Last();

            newStream = newStream.InsertRecord(sheetHeader, lastSheet8);

            List<BiffRecord> sheetRecords = RecordHelper.ParseBiffStreamBytes(sheetBytes);
            newStream = newStream.InsertRecords(sheetRecords);

            newStream = newStream.FixBoundSheetOffsets();

            return newStream;
        }

        public WorkbookStream AddSheet(BoundSheet8 sheetHeader, List<BiffRecord> records)
        {
            return AddSheet(sheetHeader, RecordHelper.ConvertBiffRecordsToBytes(records));
        }

        /// <summary>
        /// Needs to be called any time that we add a record that changes the start
        /// offset of worksheet streams. 
        /// </summary>
        /// <returns></returns>
        public WorkbookStream FixBoundSheetOffsets()
        {
            List<BoundSheet8> oldSheetBoundRecords = GetAllRecordsByType<BoundSheet8>();

            //We ignore the first BOF record for the global/workbook stream
            List<BOF> bofRecords = GetAllRecordsByType<BOF>().Skip(1).ToList();

            WorkbookStream newStream = new WorkbookStream(Records);

            int sheetOffset = 0;

            //Assign each offset in order of definition (as per specification)
            foreach (var boundSheet in oldSheetBoundRecords)
            {
                long offset = newStream.GetRecordByteOffset(bofRecords[sheetOffset]);
                BoundSheet8 newBoundSheet8 = ((BiffRecord) boundSheet.Clone()).AsRecordType<BoundSheet8>();
                newBoundSheet8.lbPlyPos = (uint)offset;
                newStream = newStream.ReplaceRecord(boundSheet, newBoundSheet8);
                sheetOffset += 1;
            }

            return newStream;
        }

        private int GetRecordOffset(BiffRecord record)
        {
            if (ContainsRecord(record) == false)
            {
                throw new ArgumentException(string.Format("Could not find record {0}", record));
            }

            var recordOffset = 
                _biffRecords.TakeWhile(r => r.Equals(record) == false).Count();

            return recordOffset;
        }


        public long GetRecordByteOffset(BiffRecord record)
        {
            int listOffset = GetRecordOffset(record);
            //Size of BiffRecord is 4 (header) + Length
            return _biffRecords.Take(listOffset).Sum(r => r.Length + 4);
        }

        public List<T> GetAllRecordsByType<T>() where T : BiffRecord
        {
            RecordType rt;
            if (RecordType.TryParse(typeof(T).Name, out rt))
            {
                return GetAllRecordsByType(rt).Select(r => (T) r.AsRecordType<T>()).ToList();
            }
            //Special edge case for the String BIFF record since it overlaps with the c# string keyword
            else if (typeof(T).Name.Equals("STRING"))
            {
                rt = RecordType.String;
                return GetAllRecordsByType(rt).Select(r => (T)r.AsRecordType<T>()).ToList();
            }

            throw new ArgumentException(string.Format("Could not identify matching RecordType for class {0}",
                typeof(T).Name));
        }

        public List<BiffRecord> GetAllRecordsByType(RecordType type)
        {
            return _biffRecords.Where(r => r.Id == type).Select(r => (BiffRecord)r.Clone()).ToList();
        }

        public List<Lbl> GetAutoOpenLabels()
        {
            List<Lbl> labels = GetAllRecordsByType<Lbl>();
            List<Lbl> autoOpenLabels = new List<Lbl>();
            foreach (var label in labels)
            {
                if (label.IsAutoOpenLabel())
                {
                    autoOpenLabels.Add(label);
                }
            }

            return autoOpenLabels;
        }

        public List<BOF> GetMacroSheetBOFs()
        {
            List<BoundSheet8> sheets = GetAllRecordsByType<BoundSheet8>().ToList();
            
            //Each BoundSheet is mapped to the 1+Nth BOF record (BoundSheet 1 is the 2nd record, etc.)
            List<BOF> bofs = GetAllRecordsByType<BOF>();

            List<BOF> macroSheetBofs = new List<BOF>();

            int sheetOffset = 1;
            foreach (var sheet in sheets)
            {
                if (sheet.dt == BoundSheet8.SheetType.Macrosheet)
                {
                    macroSheetBofs.Add(bofs[sheetOffset]);
                }

                sheetOffset += 1;
            }

            return macroSheetBofs;
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
        /// 4) The Unicode BOM character (0xFEFF/0xFFEF) is also disregarded by Excel. We can use this to break detections that will drop
        ///    nulls and look for Auto_Open without being case sensitive. By injecting this with nulls we break most detection.
        /// </summary>
        /// <returns></returns>
        public WorkbookStream ObfuscateAutoOpen(string localizedLabel = "Auto_Open")
        {
            Random randomUnicodeChar = new Random();
            string[] badUnicodeChars = { "\ufefe", "\uffff", "\ufeff", "\ufffe", "\uffef", "\ufff0", "\ufff1", "\ufff6", "\ufefd", "\u0000", "\udddd" };
            int indexLabel = 0;
            string unicodeLabelWithBadChars = "";
            List<Lbl> labels = GetAllRecordsByType<Lbl>();
            Lbl autoOpenLbl = labels.First(l => l.fBuiltin && l.Name.Value.Equals("\u0001") ||
                                                l.Name.Value.ToLower().StartsWith(localizedLabel.ToLower()));
            Lbl replaceLabelStringLbl = ((BiffRecord)autoOpenLbl.Clone()).AsRecordType<Lbl>();

            //Characters that work
            //fefe, ffff, feff, fffe, ffef, fff0, fff1, fff6, fefd, 0000, dddd
            //Pretty much any character that is invalid unicode - though \ucccc doesn't seem to work - need better criteria for parsing

            foreach (char localizedLabelChar in localizedLabel)
            {
                indexLabel = randomUnicodeChar.Next(localizedLabel.Length);
                for (var i = 0; i < 10; i += 1)
                {
                    unicodeLabelWithBadChars += badUnicodeChars[indexLabel];
                }
                unicodeLabelWithBadChars += localizedLabelChar;
            }
            replaceLabelStringLbl.SetName(new XLUnicodeStringNoCch(unicodeLabelWithBadChars, true));
            replaceLabelStringLbl.fBuiltin = false;

            // Hidden removes from the label manager entirely, but doesn't seem to work if fBuiltin is false
            // replaceLabelStringLbl.fHidden = true;

            WorkbookStream obfuscatedStream = ReplaceRecord(autoOpenLbl, replaceLabelStringLbl);
            obfuscatedStream = obfuscatedStream.FixBoundSheetOffsets();
            return obfuscatedStream;
        }

        /// <summary>
        /// Check for the existence of a FilePass BIFF record indicating RC4 or XOR Obfuscation encryption
        /// </summary>
        /// <returns>true if a FilePass record can be found</returns>
        public bool HasPasswordToOpen()
        {
            bool hasPasswordToOpen = GetAllRecordsByType<FilePass>().Count > 0;
            return hasPasswordToOpen;
        }


        public byte[] ToBytes()
        {
            return RecordHelper.ConvertBiffRecordsToBytes(_biffRecords);
        }



        public string ToDisplayString()
        {
            return string.Join("\n",_biffRecords.Select(record => record.ToHexDumpString()));
        }
    }
}
