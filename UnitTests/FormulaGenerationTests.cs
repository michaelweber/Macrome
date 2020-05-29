using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
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
        public void TestStringChunkCreation()
        {
            string chunkMe = "AAAABBBBCCCCDD";
            List<string> chunks = chunkMe.SplitStringIntoChunks(4);

            Assert.AreEqual("AAAA", chunks[0]);
            Assert.AreEqual("BBBB", chunks[1]);
            Assert.AreEqual("CCCC", chunks[2]);
            Assert.AreEqual("DD", chunks[3]);
            Assert.AreEqual(4, chunks.Count);
        }

        [Test]
        public void TestLongMacroImport()
        {
            List<String> simpleMacros = new List<string>();
            for (int cell = 0; cell < 50; cell += 1)
            {
                string cellContent = "";
                for (int i = 0; i < 0x1000; i += 1)
                {
                    cellContent += "A";
                }
                simpleMacros.Add(cellContent);
            }


            List<BiffRecord> records = FormulaHelper.ConvertStringsToRecords(simpleMacros, 0, 0xA0, 0, 0);

            BiffRecord firstGoto = records.Cast<Formula>().Last(f => f.col == 0xa0);
            BiffRecord secondColRecord1 = records.Cast<Formula>().First(f => f.col == 0xa1);

            Assert.IsNotNull(firstGoto);
            Assert.IsNotNull(secondColRecord1);
        }


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

            //The column separator character may change, so replace the old column separator with the newer internal one
            simpleMacros = simpleMacros.Select(s => s.Replace(";", MacroPatterns.MacroColumnSeparator)).ToList();

            List<BiffRecord> records = FormulaHelper.ConvertStringsToRecords(simpleMacros, 0, 0, 0, 1);

            int formulaOffset = 0;
            for (int recOffset = 3; recOffset < records.Count; recOffset += 4)
            {
                Formula f = records[recOffset].AsRecordType<Formula>();
                AbstractPtg dstPtgRef = f.ptgStack.ToList()[1];

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
