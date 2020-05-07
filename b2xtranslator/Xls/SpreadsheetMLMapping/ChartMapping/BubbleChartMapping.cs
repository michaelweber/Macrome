

using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class BubbleChartMapping : AbstractChartGroupMapping
    {
        public BubbleChartMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
            : base(workbookContext, chartContext, is3DChart)
        {
        }

        #region IMapping<CrtSequence> Members

        public override void Apply(CrtSequence crtSequence)
        {
            if (!(crtSequence.ChartType is Scatter))
            {
                throw new Exception("Invalid chart type");
            }

            var scatter = crtSequence.ChartType as Scatter;

            // c:bubbleChart 
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElBubbleChart, Dml.Chart.Ns);
            {
                // c:varyColors: This setting needs to be ignored if the chart has 
                //writeValueElement(Dml.Chart.ElVaryColors, crtSequence.ChartFormat.fVaried ? "1" : "0");

                // Bubble Chart Series
                foreach (var seriesFormatSequence in this.ChartFormatsSequence.SeriesFormatSequences)
                {
                    if (seriesFormatSequence.SerToCrt != null && seriesFormatSequence.SerToCrt.id == crtSequence.ChartFormat.idx)
                    {
                        // c:ser
                        this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSer, Dml.Chart.Ns);

                        // EG_SerShared
                        seriesFormatSequence.Convert(new SeriesMapping(this.WorkbookContext, this.ChartContext));

                        // c:dPt

                        // c:dLbls (CT_DLbls)
                        this.ChartFormatsSequence.Convert(new DataLabelMapping(this.WorkbookContext, this.ChartContext, seriesFormatSequence));

                        // c:trendline

                        // c:errBars

                        // c:xVal
                        seriesFormatSequence.Convert(new CatMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElXVal));

                        // c:yVal
                        seriesFormatSequence.Convert(new ValMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElYVal));

                        // c:bubbleSize

                        // c:bubble3D

                        this._writer.WriteEndElement(); // c:ser
                    }
                }

                // c:dLbls

                // c:bubble3D

                // c:bubbleScale

                // c:showNegBubbles

                // c:sizeRepresents


                // Axis Ids
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
