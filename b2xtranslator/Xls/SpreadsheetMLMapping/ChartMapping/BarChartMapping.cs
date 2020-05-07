

using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class BarChartMapping : AbstractChartGroupMapping
    {
        public BarChartMapping(ExcelContext workbookContext, ChartContext chartContext, bool is3DChart)
            : base(workbookContext, chartContext, is3DChart)
        {
        }

        #region IMapping<CrtSequence> Members

        public override void Apply(CrtSequence crtSequence)
        {
            if (!(crtSequence.ChartType is Bar))
            {
                throw new Exception("Invalid chart type");
            }

            var bar = crtSequence.ChartType as Bar;


            // c:barChart / c:bar3DChart
            this._writer.WriteStartElement(Dml.Chart.Prefix, this._is3DChart ? Dml.Chart.ElBar3DChart : Dml.Chart.ElBarChart, Dml.Chart.Ns);
            {
                // EG_BarChartShared
                // c:barDir
                writeValueElement(Dml.Chart.ElBarDir, bar.fTranspose ? "bar" : "col");

                // c:grouping
                string grouping = bar.fStacked ? "stacked" : bar.f100 ? "percentStacked" : this.Is3DChart && !crtSequence.Chart3d.fCluster ? "standard" : "clustered";
                writeValueElement(Dml.Chart.ElGrouping, grouping);

                // c:varyColors: This setting needs to be ignored if the chart has 
                //writeValueElement(Dml.Chart.ElVaryColors, crtSequence.ChartFormat.fVaried ? "1" : "0");

                // Bar Chart Series
                foreach (var seriesFormatSequence in this.ChartFormatsSequence.SeriesFormatSequences)
                {
                    if (seriesFormatSequence.SerToCrt != null && seriesFormatSequence.SerToCrt.id == crtSequence.ChartFormat.idx)
                    {
                        // c:ser
                        this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElSer, Dml.Chart.Ns);

                        // EG_SerShared
                        seriesFormatSequence.Convert(new SeriesMapping(this.WorkbookContext, this.ChartContext));

                        // c:invertIfNegative (stored in AreaFormat)

                        // c:pictureOptions

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

                        // c:shape (we only condider the first Chart3DBarShape found)
                        var ssSeq = seriesFormatSequence.SsSequence.Find(s => s.Chart3DBarShape != null);
                        if (ssSeq != null)
                        {
                            insertShape(ssSeq.Chart3DBarShape);
                        }

                        this._writer.WriteEndElement(); // c:ser
                    }
                }


                // Data Labels


                if (this._is3DChart)
                {
                    // c:gapWidth
                    writeValueElement(Dml.Chart.ElGapWidth, crtSequence.Chart3d.pcGap.ToString());

                    // c:gapDepth
                    writeValueElement(Dml.Chart.ElGapDepth, crtSequence.Chart3d.pcDepth.ToString());

                    // c:shape 
                    if (crtSequence.SsSequence != null && crtSequence.SsSequence.Chart3DBarShape != null)
                    {
                        insertShape(crtSequence.SsSequence.Chart3DBarShape);
                    }
                }
                else
                {
                    // c:gapWidth
                    writeValueElement(Dml.Chart.ElGapWidth, bar.pcGap.ToString());

                    // c:overlap
                    writeValueElement(Dml.Chart.ElOverlap, (-bar.pcOverlap).ToString());

                    // Series Lines


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

        private void insertShape(Chart3DBarShape chart3DBarShape)
        {
            string shape = string.Empty;
            switch (chart3DBarShape.taper)
            {
                case Chart3DBarShape.TaperType.None:
                    shape = chart3DBarShape.riser == Chart3DBarShape.RiserType.Rectangle ? "box" : "cylinder";
                    break;

                case Chart3DBarShape.TaperType.TopEach:
                    shape = chart3DBarShape.riser == Chart3DBarShape.RiserType.Rectangle ? "pyramid" : "cone";
                    break;

                case Chart3DBarShape.TaperType.TopMax:
                    shape = chart3DBarShape.riser == Chart3DBarShape.RiserType.Rectangle ? "pyramidToMax" : "coneToMax";
                    break;
            }
            if (!string.IsNullOrEmpty(shape))
            {
                writeValueElement(Dml.Chart.ElShape, shape);
            }
        }
    }
}
