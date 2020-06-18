using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.xls.XlsFileFormat;
using Macrome;
using NUnit.Framework;

namespace UnitTests
{
    [TestFixture]
    class FormulaTests
    {
        [SetUp]
        public void SetUp()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        [Test]
        public void TestA1R1C1Conversion()
        {
            string a1cell = "A1";
            string r1c1cell = "R1C1";

            Assert.AreEqual(a1cell, ExcelHelperClass.ConvertR1C1ToA1(r1c1cell));
            Assert.AreEqual(r1c1cell, ExcelHelperClass.ConvertA1ToR1C1(a1cell));

            a1cell = "C5";
            r1c1cell = "R5C3";

            Assert.AreEqual(a1cell, ExcelHelperClass.ConvertR1C1ToA1(r1c1cell));
            Assert.AreEqual(r1c1cell, ExcelHelperClass.ConvertA1ToR1C1(a1cell));

            a1cell = "FE100";
            r1c1cell = "R100C161";
            Assert.AreEqual(a1cell, ExcelHelperClass.ConvertR1C1ToA1(r1c1cell));
            Assert.AreEqual(r1c1cell, ExcelHelperClass.ConvertA1ToR1C1(a1cell));

            a1cell = "HZ46";
            r1c1cell = "R46C234";
            Assert.AreEqual(a1cell, ExcelHelperClass.ConvertR1C1ToA1(r1c1cell));
            Assert.AreEqual(r1c1cell, ExcelHelperClass.ConvertA1ToR1C1(a1cell));

            a1cell = "IU255";
            r1c1cell = "R255C255";
            Assert.AreEqual(a1cell, ExcelHelperClass.ConvertR1C1ToA1(r1c1cell));
            Assert.AreEqual(r1c1cell, ExcelHelperClass.ConvertA1ToR1C1(a1cell));

            a1cell = "AA1";
            r1c1cell = "R1C27";
            Assert.AreEqual(a1cell, ExcelHelperClass.ConvertR1C1ToA1(r1c1cell));
            Assert.AreEqual(r1c1cell, ExcelHelperClass.ConvertA1ToR1C1(a1cell));

            a1cell = "AZ1";
            r1c1cell = "R1C52";
            Assert.AreEqual(a1cell, ExcelHelperClass.ConvertR1C1ToA1(r1c1cell));
            Assert.AreEqual(r1c1cell, ExcelHelperClass.ConvertA1ToR1C1(a1cell));

        }

        [Test]
        public void TestFormulaToStringConversion()
        {
            WorkbookStream wbs = TestHelpers.GetMacroLoopWorkbookStream();

            List<RecordType> relevantTypes = new List<RecordType>()
            {
                RecordType.BoundSheet8, //Sheet definitions (Defines macro sheets + hides them)
                RecordType.Lbl,         //Named Cells (Contains Auto_Start) 
                RecordType.Formula      //The meat of most cell content
            };

            List<BiffRecord> relevantRecords = wbs.Records.Where(rec => relevantTypes.Contains(rec.Id)).ToList();
            relevantRecords = RecordHelper.ConvertToSpecificRecords(relevantRecords);

            relevantRecords = PtgHelper.UpdateGlobalsStreamReferences(relevantRecords);

            List<string> results = relevantRecords.Select(r => r.ToHexDumpString()).ToList();


            string b1formula = results.Where(res => res.StartsWith("Formula[B1]")).First();
            Assert.AreEqual("Formula[B1]: invokeChar=A11", b1formula);

            string b2formula = results.Where(res => res.StartsWith("Formula[B2]")).First();
            Assert.AreEqual("Formula[B2]: var=999", b2formula);

            string b5formula = results.Where(res => res.StartsWith("Formula[B5]")).First();
            Assert.AreEqual("Formula[B5]: InvokeFormula(\"=HALT()\",A1)", b5formula);

            string b6formula = results.Where(res => res.StartsWith("Formula[B6]")).First();
            Assert.AreEqual("Formula[B6]: WProcessMemory(-1,B2+(D1*255),ACTIVE.CELL(),LEN(ACTIVE.CELL()),0)", b6formula);

            string a11formula = results.Where(res => res.StartsWith("Formula[A11]")).First();
            Assert.AreEqual("Formula[A11]: RETURN(CHAR(var))", a11formula);

            string a12formula = results.Where(res => res.StartsWith("Formula[A12]")).First();
            Assert.AreEqual("Formula[A12]: RETURN(FORMULA(arg1,arg2))", a12formula);
            
            string d13formula = results.Where(res => res.StartsWith("Formula[D13]")).First();
            Assert.AreEqual("Formula[D13]: stringToBuild=stringToBuild&invokeChar()", d13formula);

            string d14formula = results.Where(res => res.StartsWith("Formula[D14]")).First();
            Assert.AreEqual("Formula[D14]: curCell=ABSREF(\"R[1]C\",curCell)", d14formula);
        }
    }
}
