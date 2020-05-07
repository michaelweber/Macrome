using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;

namespace b2xtranslator.SpreadsheetMLMapping
{
    /// <summary>
    /// A container class storing information required by the chart mapping classes 
    /// </summary>
    public class ChartContext
    {
        public enum ChartLocation
        {
            Embedded,
            Chartsheet
        }

        private ChartPart _chartPart;
        private ChartSheetContentSequence _chartSheetContentSequence;
        private ChartLocation _location;

        public ChartContext(ChartPart chartPart, ChartSheetContentSequence chartSheetContentSequence, ChartLocation location)
        {
            this._chartPart = chartPart;
            this._chartSheetContentSequence = chartSheetContentSequence;
            this._location = location;
        }

        public ChartPart ChartPart
        {
            get { return this._chartPart; }
        }

        public ChartSheetContentSequence ChartSheetContentSequence
        {
            get { return this._chartSheetContentSequence; }
        }
        
        public ChartLocation Location
        {
            get { return this._location; }
        }

        public ChartFormatsSequence ChartFormatsSequence
        {
            get { return this.ChartSheetContentSequence.ChartFormatsSequence; }
        }
    }
}
