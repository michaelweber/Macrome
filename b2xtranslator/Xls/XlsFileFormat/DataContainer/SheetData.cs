using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public abstract class SheetData : IVisitable
    {
        public BoundSheet8 boundsheetRecord;

        // this value is used for the case that the converter adds the 
        // chartview sheets as emty sheets
        // TODO: remove
        public bool emtpyWorksheet;

        
        public abstract void Convert<T>(T mapping);
    }
}
