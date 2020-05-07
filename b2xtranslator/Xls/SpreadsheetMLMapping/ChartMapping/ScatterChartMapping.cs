

using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class ScatterChartMapping : AbstractChartGroupMapping
    {
        public ScatterChartMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
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

            // c:scatterChart 
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElScatterChart, Dml.Chart.Ns);
            {
                // c:scatterStyle
                writeValueElement(Dml.Chart.Prefix, Dml.Chart.ElScatterStyle, Dml.Chart.Ns, mapScatterStyle(crtSequence.SsSequence));

                // c:varyColors
                //writeValueElement(Dml.Chart.ElVaryColors, crtSequence.ChartFormat.fVaried ? "1" : "0");

                foreach (var seriesFormatSequence in this.ChartFormatsSequence.SeriesFormatSequences)
                {
                    if (seriesFormatSequence.SerToCrt != null && seriesFormatSequence.SerToCrt.id == crtSequence.ChartFormat.idx)
                    {
                        // c:ser (CT_ScatterSer)
                        // c:ser
                        this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSer, Dml.Chart.Ns);

                        // EG_SerShared
                        seriesFormatSequence.Convert(new SeriesMapping(this.WorkbookContext, this.ChartContext));

                        // c:marker

                        // c:dPt
                        
                        // c:dLbls (CT_DLbls)
                        this.ChartFormatsSequence.Convert(new DataLabelMapping(this.WorkbookContext, this.ChartContext, seriesFormatSequence));

                        // c:trendline

                        // c:errBars

                        // c:xVal
                        seriesFormatSequence.Convert(new CatMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElXVal));

                        // c:yVal
                        seriesFormatSequence.Convert(new ValMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElYVal));

                        // c:smooth
                        writeValueElement(Dml.Chart.Prefix, Dml.Chart.ElSmooth, Dml.Chart.Ns, isSmoothed(crtSequence.SsSequence) ? "1" : "0");

                        this._writer.WriteEndElement(); // c:ser
                    }
                }

                // Data Labels

                // Axis Ids
                foreach (int axisId in crtSequence.ChartFormat.AxisIds)
                {
                    writeValueElement(Dml.Chart.ElAxId, axisId.ToString());
                }
            }
            this._writer.WriteEndElement();
        }
        #endregion

        private string mapScatterStyle(SsSequence ssSequence)
        {
            // CT_ScatterStyle
            // The following scatter styles exist: line, lineMarker, marker, none, smooth, smoothMarker
            // 
            bool smoothed = isSmoothed(ssSequence);
            bool hasMarker = (ssSequence == null) || (ssSequence.MarkerFormat != null && ssSequence.MarkerFormat.imk != MarkerFormat.MarkerType.NoMarker);
            bool hasLine = (ssSequence == null) || (ssSequence.LineFormat != null && ssSequence.LineFormat.lns != LineFormat.LineStyle.None);

            string scatterStyle = "none";
            if (smoothed && hasMarker)
            {
                scatterStyle = "smoothMarker";
            }
            else if (smoothed)
            {
                scatterStyle = "smooth";
            }
            else if (hasMarker && hasLine)
            {
                scatterStyle = "lineMarker";
            }
            else if (hasLine)
            {
                scatterStyle = "line";
            }
            else if (hasMarker)
            {
                scatterStyle = "marker";
            }
            return scatterStyle;
        }

        private bool isSmoothed(SsSequence ssSequence)
        {
            return (ssSequence != null && ssSequence.SerFmt != null && ssSequence.SerFmt.fSmoothedLine);
        }
    }
}
