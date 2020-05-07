

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public abstract class AbstractChartMapping : AbstractOpenXmlMapping
    {
        private ExcelContext _workbookContext;
        private ChartContext _chartContext;
        
        public AbstractChartMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(chartContext.ChartPart.XmlWriter)
        {
            this._workbookContext = workbookContext;
            this._chartContext = chartContext;
        }

        public ExcelContext WorkbookContext
        {
            get { return this._workbookContext; }
        }

        public ChartContext ChartContext
        {
            get { return this._chartContext; }
        }

        public ChartPart ChartPart
        {
            get { return this.ChartContext.ChartPart; }
        }

        public ChartSheetContentSequence ChartSheetContentSequence
        {
            get { return this.ChartContext.ChartSheetContentSequence; }
        }

        public ChartContext.ChartLocation Location
        {
            get { return this.ChartContext.Location; }
        }

        public ChartFormatsSequence ChartFormatsSequence
        {
            get { return this.ChartSheetContentSequence.ChartFormatsSequence; }
        }

        protected void writeValueElement(string localName, string value)
        {
            this._writer.WriteStartElement(Dml.Chart.Prefix, localName, Dml.Chart.Ns);
            this._writer.WriteAttributeString(Dml.BaseTypes.AttrVal, value);
            this._writer.WriteEndElement();
        }

        protected void writeValueElement(string prefix, string localName, string ns, string value)
        {
            this._writer.WriteStartElement(prefix, localName, ns);
            this._writer.WriteAttributeString(Dml.BaseTypes.AttrVal, value);
            this._writer.WriteEndElement();
        }
    }
}
