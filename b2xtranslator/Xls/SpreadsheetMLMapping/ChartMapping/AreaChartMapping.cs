

using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class AreaChartMapping : AbstractChartGroupMapping
    {
        public AreaChartMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
            : base(workbookContext, chartContext, is3DChart)
        {
        }

        #region IMapping<CrtSequence> Members

        public override void Apply(CrtSequence crtSequence)
        {
            if (!(crtSequence.ChartType is Area))
            {
                throw new Exception("Invalid chart type");
            }

            var area = crtSequence.ChartType as Area;

            // c:areaChart / c:area3DChart
            this._writer.WriteStartElement(Dml.Chart.Prefix, this.Is3DChart ? Dml.Chart.ElArea3DChart : Dml.Chart.ElAreaChart, Dml.Chart.Ns);
            {
                // EG_AreaChartShared
                // CT_Grouping

                // c:varyColors: This setting needs to be ignored if the chart has 
                //writeValueElement(Dml.Chart.ElVaryColors, crtSequence.ChartFormat.fVaried ? "1" : "0");

                // Area Chart Series (CT_AreaSer)
                foreach (var seriesFormatSequence in this.ChartFormatsSequence.SeriesFormatSequences)
                {
                    if (seriesFormatSequence.SerToCrt != null && seriesFormatSequence.SerToCrt.id == crtSequence.ChartFormat.idx)
                    {
                        // c:ser
                        this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSer, Dml.Chart.Ns);

                        // EG_SerShared
                        seriesFormatSequence.Convert(new SeriesMapping(this.WorkbookContext, this.ChartContext));

                        // c:pictureOptions (CT_PictureOptions)
                        
                        // c:dPt (Data Points)
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

                        // c:cat (Category Axis Data)
                        seriesFormatSequence.Convert(new CatMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElCat));

                        // c:val
                        seriesFormatSequence.Convert(new ValMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElVal));

                        this._writer.WriteEndElement(); // c:ser
                    }
                }
        
                // c:dLbls (Data Labels)

                // c:dropLines


                if (this.Is3DChart)
                {
                    // c:gapDepth
                }

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
