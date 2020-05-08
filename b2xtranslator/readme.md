# b2xtranslator 

The XLS BIFF record parsing has been taken from the `b2xtranslator` project at https://github.com/EvolutionJobs/b2xtranslator/. This code base is meant for parsing arbitrary legacy 2003 XLS files and converting them to the newer XML format. Each record is assumed to be associated with a file on disk so records all have a stream and offset to access their raw bytes. This works for modifying existing records we read in, but if you want to create records from scratch it doesn't provide the needed functionality. 

The code base here has been modified to allow manual creation and modification of BIFF records and PTG records used in Excel macros. Changes have largely been made to the `BiffRecord` and `AbstractPtg` classes with some additional overrides added for relevant record types such as:

* BOF
* BoundSheet8
* Lbl
* ExternSheet
* Formula
* Lbl
* STRING
* SupBook

A more hacky approach was taken for Ptg (Parse Thing) records, and most of the binary writing code resides in `PtgHelper.cs` because just appending `Helper` to names is a GREAT development practice. 

All code retained from that version ©2009 DI<sup><u>a</u></sup>LOGIK<sup><u>a</u></sup> http://www.dialogika.de/  
.NET core port work and move to `System.IO.Compression` ©2017 Evolution https://www.evolutionjobs.com/
