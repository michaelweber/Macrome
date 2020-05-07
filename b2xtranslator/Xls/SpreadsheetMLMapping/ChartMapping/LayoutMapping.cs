

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System.Globalization;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class LayoutMapping : AbstractChartMapping,
          IMapping<CrtLayout12>
    {
        public LayoutMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        #region IMapping<CrtLayout12> Members

        public void Apply(CrtLayout12 crtLayout12)
        {
            // c:layout
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElLayout, Dml.Chart.Ns);
            {
                if (crtLayout12.wHeightMode != CrtLayout12.CrtLayout12Mode.L12MAUTO ||
                    crtLayout12.wWidthMode != CrtLayout12.CrtLayout12Mode.L12MAUTO ||
                    crtLayout12.wYMode != CrtLayout12.CrtLayout12Mode.L12MAUTO ||
                    crtLayout12.wXMode != CrtLayout12.CrtLayout12Mode.L12MAUTO)
                {
                    // c:manualLayout
                    this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElManualLayout, Dml.Chart.Ns);
                    {
                        // c:layoutTarget
                        writeValueElement(Dml.Chart.ElLayoutTarget, crtLayout12.fLayoutTargetInner ? "inner" : "outer");

                        if (crtLayout12.wXMode != CrtLayout12.CrtLayout12Mode.L12MAUTO)
                        {
                            // c:xMode
                            writeValueElement(Dml.Chart.ElXMode, crtLayout12.wXMode == CrtLayout12.CrtLayout12Mode.L12MEDGE ? "edge" : "factor");
                        }
                        if (crtLayout12.wYMode != CrtLayout12.CrtLayout12Mode.L12MAUTO)
                        {
                            // c:yMode
                            writeValueElement(Dml.Chart.ElYMode, crtLayout12.wYMode == CrtLayout12.CrtLayout12Mode.L12MEDGE ? "edge" : "factor");
                        }
                        if (crtLayout12.wWidthMode != CrtLayout12.CrtLayout12Mode.L12MAUTO)
                        {
                            // c:wMode
                            writeValueElement(Dml.Chart.ElWMode, crtLayout12.wWidthMode == CrtLayout12.CrtLayout12Mode.L12MEDGE ? "edge" : "factor");
                        }
                        if (crtLayout12.wHeightMode != CrtLayout12.CrtLayout12Mode.L12MAUTO)
                        {
                            // c:hMode
                            writeValueElement(Dml.Chart.ElHMode, crtLayout12.wHeightMode == CrtLayout12.CrtLayout12Mode.L12MEDGE ? "edge" : "factor");
                        }

                        // c:x
                        writeValueElement(Dml.Chart.ElX, crtLayout12.x.ToString(CultureInfo.InvariantCulture));

                        // c:y
                        writeValueElement(Dml.Chart.ElY, crtLayout12.y.ToString(CultureInfo.InvariantCulture));

                        // c:w
                        writeValueElement(Dml.Chart.ElW, crtLayout12.dx.ToString(CultureInfo.InvariantCulture));

                        // c:h
                        writeValueElement(Dml.Chart.ElH, crtLayout12.dy.ToString(CultureInfo.InvariantCulture));
                    }
                    this._writer.WriteEndElement(); // c:manualLayout
                }
            }
            this._writer.WriteEndElement(); // c:layout
        }
        #endregion
    }
}
