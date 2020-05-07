

using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class LineChartMapping : AbstractChartGroupMapping
    {
        public LineChartMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
            : base(workbookContext, chartContext, is3DChart)
        {
        }

        #region IMapping<CrtSequence> Members

        public override void Apply(CrtSequence crtSequence)
        {
            if (!(crtSequence.ChartType is Line))
            {
                throw new Exception("Invalid chart type");
            }

            var line = crtSequence.ChartType as Line;

            // c:lineChart or c:stockChart 
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElLineChart, Dml.Chart.Ns);
            {
                // EG_LineChartShared
                // c:grouping
                string grouping = line.fStacked ? "stacked" : line.f100 ? "percentStacked" : "standard";
                writeValueElement(Dml.Chart.ElGrouping, grouping);

                // c:varyColors
                writeValueElement(Dml.Chart.ElVaryColors, crtSequence.ChartFormat.fVaried ? "1" : "0");

                // Line Chart Series
                foreach (var seriesFormatSequence in this.ChartFormatsSequence.SeriesFormatSequences)
                {
                    if (seriesFormatSequence.SerToCrt != null && seriesFormatSequence.SerToCrt.id == crtSequence.ChartFormat.idx)
                    {
                        // c:ser
                        this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSer, Dml.Chart.Ns);

                        // EG_SerShared
                        seriesFormatSequence.Convert(new SeriesMapping(this.WorkbookContext, this.ChartContext));

                        // c:marker

                        // c:dPt
                        for (int i = 1; i < seriesFormatSequence.SsSequence.Count; i++)
                        {
                            // write a dPt for each SsSequence
                            var ssSequence = seriesFormatSequence.SsSequence[i];
                            ssSequence.Convert(new DataPointMapping(this.WorkbookContext, this.ChartContext, i - 1));
                        }

                        // c:dLbls (Data Labels)
                        this.ChartFormatsSequence.Convert(new DataLabelMapping(this.WorkbookContext, this.ChartContext, seriesFormatSequence));

                        // c:trendline

                        // c:errBars

                        // c:cat

                        // c:val
                        seriesFormatSequence.Convert(new ValMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElVal));

                        // c:smooth

                        // c:shape

                        this._writer.WriteEndElement(); // c:ser
                    }
                }
                
                // c:dLbls

                // dropLines

                // End EG_LineChartShared

                if (this.Is3DChart)
                {
                    // c:gapDepth
                }
                else
                {
                    // c:hiLowLines

                    // c:upDownBars

                    // c:marker

                    // c:smooth

                }

                // c:axId
                foreach (int axisId in crtSequence.ChartFormat.AxisIds)
                {
                    writeValueElement(Dml.Chart.ElAxId, axisId.ToString());
                }
            }
            this._writer.WriteEndElement();
        }
        #endregion
    }
}
