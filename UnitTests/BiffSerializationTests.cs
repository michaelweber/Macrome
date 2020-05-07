using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.Tools;
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

    }
}
