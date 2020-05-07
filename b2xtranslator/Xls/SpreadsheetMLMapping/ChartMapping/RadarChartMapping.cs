

using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class RadarChartMapping : AbstractChartGroupMapping
    {
        public RadarChartMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
            : base(workbookContext, chartContext, is3DChart)
        {
        }

        #region IMapping<CrtSequence> Members

        public override void Apply(CrtSequence crtSequence)
        {
            if (!(crtSequence.ChartType is Radar))
            {
                throw new Exception("Invalid chart type");
            }

            var radar = crtSequence.ChartType as Radar;

            // c:radarChart
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElRadarChart, Dml.Chart.Ns);
            {
                // c:radarStyle
                writeValueElement(Dml.Chart.ElRadarStyle, mapRadarStyle(crtSequence.SsSequence));

                // c:varyColors: This setting needs to be ignored if the chart has 
                //writeValueElement(Dml.Chart.ElVaryColors, crtSequence.ChartFormat.fVaried ? "1" : "0");

                // Radar Chart Series (CT_RadarSer)
                foreach (var seriesFormatSequence in this.ChartFormatsSequence.SeriesFormatSequences)
                {
                    if (seriesFormatSequence.SerToCrt != null && seriesFormatSequence.SerToCrt.id == crtSequence.ChartFormat.idx)
                    {
                        // c:ser
                        this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSer, Dml.Chart.Ns);

                        // EG_SerShared
                        seriesFormatSequence.Convert(new SeriesMapping(this.WorkbookContext, this.ChartContext));

                        // c:marker

                        // c:dPt (Data Points)
                        for (int i = 1; i < seriesFormatSequence.SsSequence.Count; i++)
                        {
                            // write a dPt for each SsSequence
                            var sss = seriesFormatSequence.SsSequence[i];
                            sss.Convert(new DataPointMapping(this.WorkbookContext, this.ChartContext, i - 1));
                        }

                        // c:dLbls (Data Labels)
                        this.ChartFormatsSequence.Convert(new DataLabelMapping(this.WorkbookContext, this.ChartContext, seriesFormatSequence));

                        // c:cat (Category Axis Data)
                        seriesFormatSequence.Convert(new CatMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElCat));

                        // c:val
                        seriesFormatSequence.Convert(new ValMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElVal));

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

        private string mapRadarStyle(SsSequence ssSequence)
        {
            string radarStyle = "standard";
            bool hasMarker = (ssSequence == null) || (ssSequence.MarkerFormat != null && ssSequence.MarkerFormat.imk != MarkerFormat.MarkerType.NoMarker);

            if (hasMarker)
            {
                radarStyle = "marker";
            }
            // TODO: radar fill style

            return radarStyle;
        }
    }
}
