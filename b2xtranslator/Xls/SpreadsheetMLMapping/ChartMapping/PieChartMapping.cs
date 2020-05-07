

using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class PieChartMapping : AbstractChartGroupMapping
    {
        public PieChartMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
            : base(workbookContext, chartContext, is3DChart)
        {
        }

        #region IMapping<CrtSequence> Members

        public override void Apply(CrtSequence crtSequence)
        {
            if (!(crtSequence.ChartType is Pie))
            {
                throw new Exception("Invalid chart type");
            }

            var pie = crtSequence.ChartType as Pie;

            bool isDoughnutChart = (pie.pcDonut != 0);
            
            string chartType = this._is3DChart ? Dml.Chart.ElPie3DChart : Dml.Chart.ElPieChart; 
            if (isDoughnutChart)
            {
                chartType = Dml.Chart.ElDoughnutChart;
            }

            // c:pieChart or c:pie3DChart or c:doughnutChart
            this._writer.WriteStartElement(Dml.Chart.Prefix, chartType, Dml.Chart.Ns);
            {
                // EG_PieChartShared

                // varyColors
                writeValueElement("varyColors", crtSequence.ChartFormat.fVaried ? "1" : "0");

                // Pie Chart Series (CT_PieSer)
                foreach (var seriesFormatSequence in this.ChartFormatsSequence.SeriesFormatSequences)
                {
                    if (seriesFormatSequence.SerToCrt != null && seriesFormatSequence.SerToCrt.id == crtSequence.ChartFormat.idx)
                    {
                        // c:ser
                        this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSer, Dml.Chart.Ns);

                        // EG_SerShared
                        seriesFormatSequence.Convert(new SeriesMapping(this.WorkbookContext, this.ChartContext));

                        // c:explosion
                        var sssBase = seriesFormatSequence.SsSequence[0];
                        if (sssBase.PieFormat != null)
                        {
                            writeValueElement("explosion", sssBase.PieFormat.pcExplode.ToString());
                        }

                        // c:dPt (Data Points)
                        for (int i = 1; i < seriesFormatSequence.SsSequence.Count; i++)
                        {
                            // write a dPt for each SsSequence
                            var sss = seriesFormatSequence.SsSequence[i];
                            sss.Convert(new DataPointMapping(this.WorkbookContext, this.ChartContext, i-1));
                        }

                        // c:dLbls (Data Labels)
                        this.ChartFormatsSequence.Convert(new DataLabelMapping(this.WorkbookContext, this.ChartContext, seriesFormatSequence));

                        // c:cat
                        seriesFormatSequence.Convert(new CatMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElCat));

                        // c:val
                        seriesFormatSequence.Convert(new ValMapping(this.WorkbookContext, this.ChartContext, Dml.Chart.ElVal));

                        this._writer.WriteEndElement(); // c:ser
                    }
                }

                if (!this.Is3DChart)
                {
                    // c:firstSliceAng
                    writeValueElement("firstSliceAng", pie.anStart.ToString());
                }
                if (isDoughnutChart)
                {
                    // c:holeSize
                }
            }
            this._writer.WriteEndElement();
        }
        #endregion
    }
}
