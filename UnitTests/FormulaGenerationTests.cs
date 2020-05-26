using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using Macrome;
using NUnit.Framework;

namespace UnitTests
{
    [TestFixture]
    class FormulaGenerationTests
    {
        [Test]
        public void TestMultiColumnMacroImport()
        {
            List<String> simpleMacros = new List<string>()
            {
                "A1;B1;C1",
                "A2;B2;C2",
                "A3;B3",
                "A4"
            };

            List<BiffRecord> records = FormulaHelper.ConvertStringsToRecords(simpleMacros, 0, 0, 0, 1);

            int formulaOffset = 0;
            for (int recOffset = 3; recOffset < records.Count; recOffset += 4)
            {
                Formula f = records[recOffset].AsRecordType<Formula>();
                AbstractPtg formulaFunctionPtg = f.ptgStack.ToList()[0];
                AbstractPtg dstPtgRef = f.ptgStack.ToList()[1];
                AbstractPtg srcPtgRef = f.ptgStack.ToList()[2];

                PtgRef dst = (PtgRef) dstPtgRef;
                int targetRow = dst.rw;
                int targetCol = dst.col & 0xFF;

                switch(formulaOffset)
                {
                    case 0:
                        Assert.AreEqual(0, targetRow);
                        Assert.AreEqual(1, targetCol);
                        break;
                    case 1:
                        Assert.AreEqual(0, targetRow);
                        Assert.AreEqual(2, targetCol);
                        break;
                    case 2:
                        Assert.AreEqual(0, targetRow);
                        Assert.AreEqual(3, targetCol);
                        break;
                    case 3:
                        Assert.AreEqual(1, targetRow);
                        Assert.AreEqual(1, targetCol);
                        break;
                    case 4:
                        Assert.AreEqual(1, targetRow);
                        Assert.AreEqual(2, targetCol);
                        break;
                    case 5:
                        Assert.AreEqual(1, targetRow);
                        Assert.AreEqual(3, targetCol);
                        break;
                    case 6:
                        Assert.AreEqual(2, targetRow);
                        Assert.AreEqual(1, targetCol);
                        break;
                    case 7:
                        Assert.AreEqual(2, targetRow);
                        Assert.AreEqual(2, targetCol);
                        break;
                    case 8:
                        Assert.AreEqual(3, targetRow);
                        Assert.AreEqual(1, targetCol);
                        break;
                    default:
                        Assert.Fail("Too many records generated");
                        break;
                }

                formulaOffset += 1;
            }
        }
    }
}
