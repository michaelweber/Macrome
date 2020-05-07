using b2xtranslator.CommonTranslatorLib;
namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class ChartSheetData : SheetData
    {
        public ChartSheetSequence ChartSheetSequence;

        public ChartSheetData()
        {
        }

        public override void Convert<T>(T mapping)
        {
            ((IMapping<ChartSheetSequence>)mapping).Apply(this.ChartSheetSequence);
        }
    }
}
