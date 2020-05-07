

using System;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.StructuredStorage.Writer;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class MacroBinaryMapping : AbstractOpenXmlMapping,
        IMapping<XlsDocument>
    {
        private ExcelContext ctx;
        private string projectFolder = "\\_VBA_PROJECT_CUR";
        private string vbaFolder = "\\_VBA_PROJECT_CUR\\VBA";
        private string projectFile = "\\_VBA_PROJECT_CUR\\PROJECT";
        private string projectWmFile = "\\_VBA_PROJECT_CUR\\PROJECTwm";

        public MacroBinaryMapping(ExcelContext ctx)
            : base(null)
        {
            this.ctx = ctx;
        }

        public void Apply(XlsDocument xls)
        {
            //get the Class IDs of the directories
            var macroClsid = new Guid();
            var vbaClsid = new Guid();
            foreach (var entry in xls.Storage.AllEntries)
            {
                if (entry.Path == this.projectFolder)
                {
                    macroClsid = entry.ClsId;
                }
                else if (entry.Path == this.vbaFolder)
                {
                    vbaClsid = entry.ClsId;
                }
            }

            //create a new storage
            var storage = new StructuredStorageWriter();
            storage.RootDirectoryEntry.setClsId(macroClsid);

            //copy the VBA directory
            var vba = storage.RootDirectoryEntry.AddStorageDirectoryEntry("VBA");
            vba.setClsId(vbaClsid);
            foreach (var entry in xls.Storage.AllStreamEntries)
            {
                if (entry.Path.StartsWith(this.vbaFolder))
                {
                    vba.AddStreamDirectoryEntry(entry.Name, xls.Storage.GetStream(entry.Path));
                }
            }

            //copy the project streams
            storage.RootDirectoryEntry.AddStreamDirectoryEntry("PROJECT", xls.Storage.GetStream(this.projectFile));
            storage.RootDirectoryEntry.AddStreamDirectoryEntry("PROJECTwm", xls.Storage.GetStream(this.projectWmFile));

           //write the storage to the xml part
            storage.write(this.ctx.SpreadDoc.WorkbookPart.VbaProjectPart.GetStream());
        }
    }
}
