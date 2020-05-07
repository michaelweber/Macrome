

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public abstract class AbstractAxisMapping : AbstractChartMapping,
          IMapping<AxesSequence>
    {
        public AbstractAxisMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        #region IMapping<AxesSequence> Members

        public virtual void Apply(AxesSequence axesSequence)
        {
            // EG_AxShared
            writeValueElement(Dml.Chart.ElAxId, axesSequence.IvAxisSequence.Axis.AxisId.ToString());
        }

        #endregion
    }
}
