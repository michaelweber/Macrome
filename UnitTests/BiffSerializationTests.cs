using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Text;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using b2xtranslator.xls.XlsFileFormat;
using b2xtranslator.xls.XlsFileFormat.Records;
using b2xtranslator.xls.XlsFileFormat.Structures;
using Macrome;
using NUnit.Framework;

namespace UnitTests
{
    [TestFixture]
    class BiffSerializationTests
    {
        [SetUp]
        public void SetUp()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        [Test]
        public void TestConstructorAndToBytes()
        {
            byte[] wbBytes = TestHelpers.GetTemplateMacroBytes();
            WorkbookStream wbs = new WorkbookStream(wbBytes);

            byte[] parsedBytes = wbs.ToBytes();

            Assert.AreEqual(wbBytes.Length, parsedBytes.Length);
            Assert.AreEqual(wbBytes, parsedBytes);
        }

        [Test]
        public void TestParseMassSelectionWithArgument()
        {
            WorkbookStream wbs = TestHelpers.GetMassSelectUDFArgumentSheet();

            string dumpString = RecordHelper.GetRelevantRecordDumpString(wbs);

            Console.WriteLine(dumpString);
        }

        [Test]
        public void TestBiffRecordCloneAndGetBytes()
        {
            BoundSheet8 bs8 = new BoundSheet8(BoundSheet8.HiddenState.Visible, BoundSheet8.SheetType.Worksheet, "MySheetName");
            byte[] bs8bytes = bs8.GetBytes();
            BiffRecord bs8CloneRecord = (BiffRecord)bs8.Clone();
            byte[] bs8cloneBytes = bs8CloneRecord.GetBytes();
            Assert.AreEqual(bs8bytes, bs8cloneBytes);
            BoundSheet8 bs8Clone = bs8CloneRecord.AsRecordType<BoundSheet8>();
            byte[] bs8asRecordBytes = bs8Clone.GetBytes();
            Assert.AreEqual(bs8bytes, bs8asRecordBytes);
        }

        [Test]
        public void TestFormulaCloneAndGetBytes()
        {
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>();
            ptgStack.Push(new PtgFunc(FtabValues.CHAR));
            ptgStack.Push(new PtgInt(42, AbstractPtg.PtgDataType.VALUE));
            Formula f = new Formula(new Cell(0, 0, 0), FormulaValue.GetStringFormulaValue(), true, new CellParsedFormula(ptgStack));

            Formula fClone = ((BiffRecord) f.Clone()).AsRecordType<Formula>();

            byte[] fBytes = f.GetBytes();
            byte[] fCloneBytes = fClone.GetBytes();

            Assert.AreEqual(28, f.Length);
            Assert.AreEqual(32, fBytes.Length);

            Assert.AreEqual(f.Length, fClone.Length);
            Assert.AreEqual(fBytes.Length, fCloneBytes.Length);
            Assert.AreEqual(fBytes, fCloneBytes);
        }

        [Test]
        public void CharRoundMacroFormulaTest()
        {
            WorkbookStream wbs = TestHelpers.GetCharRoundMacroWorkbookStream();
            List<Formula> formulas = wbs.GetAllRecordsByType<Formula>();
            Formula firstFormula = formulas.First();

            List<AbstractPtg> ptgStack = firstFormula.ptgStack.Reverse().ToList();
            Assert.AreEqual(typeof(PtgInt),ptgStack[0].GetType());
            Assert.AreEqual(typeof(PtgInt), ptgStack[1].GetType());
            Assert.AreEqual(typeof(PtgFunc), ptgStack[2].GetType());
            Assert.AreEqual(typeof(PtgFunc), ptgStack[3].GetType());
        }

        [Test]
        public void TestLabelSerialization()
        {
            WorkbookStream macroWorkbookStream = TestHelpers.GetMultiSheetMacroBytes();

            List<SupBook> supBooks = macroWorkbookStream.GetAllRecordsByType<SupBook>();
            List<Lbl> labels = macroWorkbookStream.GetAllRecordsByType<Lbl>();

            Lbl lastLabel = labels.Last();
            Lbl cloneLabel = ((BiffRecord)lastLabel.Clone()).AsRecordType<Lbl>();
            byte[] labelBytes = lastLabel.GetBytes();
            byte[] cloneBytes = cloneLabel.GetBytes();

            Assert.AreEqual(labelBytes, cloneBytes);
        }

        [Test]
        public void AutoOpenSerializationTest()
        {
            WorkbookStream autoOpenStream = new WorkbookStream(TestHelpers.GetAutoOpenTestBytes());
            List<Lbl> lbls = autoOpenStream.GetAllRecordsByType<Lbl>();

            //auto_open_test.xls contains a Lbl for Auto_Open222
            Lbl autoOpenLabel = lbls.First(l => l.fBuiltin);
            byte[] labelBytes = autoOpenLabel.Name.Bytes;

            //Should be length 4, 1 byte for the builtin string lookup, 3 bytes for 222
            Assert.AreEqual(4, autoOpenLabel.cch);

            //Not a unicode string, so fHighBit is 0
            Assert.AreEqual((byte)0x00, labelBytes[0]);
            //Starts with the Auto_Open builtin value of 1
            Assert.AreEqual((byte)0x01, labelBytes[1]);

            //Should be followed by whatever we append to the end, in this case 222
            Assert.AreEqual((byte)'2', labelBytes[2]);
            Assert.AreEqual((byte)'2', labelBytes[3]);
            Assert.AreEqual((byte)'2', labelBytes[4]);
        }

        [Test]
        public void ByteToBitmaskTest()
        {
            byte val = (byte)7;

            uint mask = Utils.ByteToBitmask(val, 0x0003C000);
            byte output = Utils.BitmaskToByte(mask, 0x0003C000);

            Assert.AreEqual(val, output);

            uint mask2 = Utils.ByteToBitmask(val, 0x00FF);
            byte output2 = Utils.BitmaskToByte(mask2, 0x00FF);

            Assert.AreEqual(val, output2);
        }


        [Test]
        public void TestGetDefaultMacroSheetInternationalized()
        {

            Intl manuallyCreatedIntlRecord = new Intl();

            byte[] intlBytes = manuallyCreatedIntlRecord.GetBytes();

            BiffRecord rec = new BiffRecord(intlBytes);

            Intl convertedRecord = rec.AsRecordType<Intl>();

            Assert.AreEqual(convertedRecord.GetBytes(), manuallyCreatedIntlRecord.GetBytes());

            WorkbookStream wbs = TestHelpers.GetDefaultMacroTemplate();
            List<BiffRecord> sheetRecords = wbs.GetRecordsForBOFRecord(wbs.GetAllRecordsByType<BOF>().Last());

            WorkbookStream internationalWbs = new WorkbookStream(sheetRecords);
            var intlRecord = new Intl();
            internationalWbs = internationalWbs.InsertRecord(intlRecord, internationalWbs.GetAllRecordsByType<b2xtranslator.Spreadsheet.XlsFileFormat.Records.Index>().First());
            Assert.IsTrue(internationalWbs.ContainsRecord(intlRecord));
            var nextRecord = internationalWbs.Records.SkipWhile(r => r.Id != RecordType.Intl).Skip(1).Take(1).First();
            Assert.IsTrue(nextRecord.Id == RecordType.CalcMode);
        }

        [Test]
        public void TestParseRC4CryptoAPIFilePassRecord()
        {
            byte[] rc4CryptoApiFilePassRecord = new byte[]
            {
                0x2F, 0x00, 0xC8, 0x00, 0x01, 0x00, 0x04, 0x00, 0x02, 0x00, 0x0C, 0x00, 0x00, 0x00, 0x7E, 0x00, 0x00,
                0x00, 0x0C, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01, 0x68, 0x00, 0x00, 0x04, 0x80, 0x00, 0x00,
                0x80, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x4D,
                0x00, 0x69, 0x00, 0x63, 0x00, 0x72, 0x00, 0x6F, 0x00, 0x73, 0x00, 0x6F, 0x00, 0x66, 0x00, 0x74, 0x00,
                0x20, 0x00, 0x45, 0x00, 0x6E, 0x00, 0x68, 0x00, 0x61, 0x00, 0x6E, 0x00, 0x63, 0x00, 0x65, 0x00, 0x64,
                0x00, 0x20, 0x00, 0x43, 0x00, 0x72, 0x00, 0x79, 0x00, 0x70, 0x00, 0x74, 0x00, 0x6F, 0x00, 0x67, 0x00,
                0x72, 0x00, 0x61, 0x00, 0x70, 0x00, 0x68, 0x00, 0x69, 0x00, 0x63, 0x00, 0x20, 0x00, 0x50, 0x00, 0x72,
                0x00, 0x6F, 0x00, 0x76, 0x00, 0x69, 0x00, 0x64, 0x00, 0x65, 0x00, 0x72, 0x00, 0x20, 0x00, 0x76, 0x00,
                0x31, 0x00, 0x2E, 0x00, 0x30, 0x00, 0x00, 0x00, 0x10, 0x00, 0x00, 0x00, 0x98, 0xC3, 0xA1, 0xA0, 0x9A,
                0xC8, 0x92, 0x3D, 0x7A, 0x5D, 0x03, 0x30, 0x34, 0xE8, 0x67, 0x64, 0xA8, 0xA9, 0x29, 0xF0, 0xA6, 0xA8,
                0xB6, 0xA1, 0x1F, 0x8C, 0x6F, 0x27, 0xB6, 0xA8, 0x82, 0xEB, 0x14, 0x00, 0x00, 0x00, 0xE8, 0xAF, 0x8B,
                0x67, 0x4E, 0x3F, 0xA7, 0xA1, 0x39, 0x9D, 0x9F, 0x2D, 0xDC, 0xB5, 0x1A, 0x33, 0x5D, 0x8A, 0xB2, 0x69
            };

            VirtualStreamReader vsr = new VirtualStreamReader(new MemoryStream(rc4CryptoApiFilePassRecord));
            FilePass filePassRecord = new FilePass(vsr, (RecordType) vsr.ReadUInt16(), vsr.ReadUInt16());

            byte[] filePassBytes = filePassRecord.GetBytes();

            Assert.AreEqual(rc4CryptoApiFilePassRecord.Length, filePassBytes.Length);

            Assert.IsFalse(filePassRecord.encryptionHeader.fDocProps);
            Assert.IsFalse(filePassRecord.encryptionHeader.fCryptoAPI);
            Assert.IsTrue(filePassRecord.encryptionHeader.fExternal);
            Assert.IsTrue(filePassRecord.encryptionHeader.fAES);

            for (int offset = 0; offset < rc4CryptoApiFilePassRecord.Length; offset += 1)
            {
                Assert.AreEqual(rc4CryptoApiFilePassRecord[offset], filePassBytes[offset]);
            }
        }
    }
}
