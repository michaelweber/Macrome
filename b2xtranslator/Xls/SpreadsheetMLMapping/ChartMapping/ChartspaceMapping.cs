

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class ChartMapping : AbstractChartMapping,
          IMapping<ChartSheetContentSequence>
    {
        ChartContext _chartContext;

        public ChartMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
            this._chartContext = chartContext;
        }

        #region IMapping<ChartSheetContentSequence> Members

        // TODO: maybe we only need chartSheetContentSequence.ChartFormatsSequence here
        public void Apply(ChartSheetContentSequence chartSheetContentSequence)
        {
            var chartFormatsSequence = chartSheetContentSequence.ChartFormatsSequence;

            // c:chartspace
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElChartSpace, Dml.Chart.Ns);
            this._writer.WriteAttributeString("xmlns", Dml.Chart.Prefix, "", Dml.Chart.Ns);
            this._writer.WriteAttributeString("xmlns", Dml.Prefix, "", Dml.Ns);
            {
                // c:chart
                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElChart, Dml.Chart.Ns);
                {
                    // c:title
                    foreach (var attachedLabelSequence in chartFormatsSequence.AttachedLabelSequences)
                    {
                        if (attachedLabelSequence.ObjectLink != null && attachedLabelSequence.ObjectLink.wLinkObj == ObjectLink.ObjectType.Chart)
                        {
                            attachedLabelSequence.Convert(new TitleMapping(this.WorkbookContext, this.ChartContext));
                            break;
                        }
                    }

                    // c:plotArea
                    chartFormatsSequence.Convert(new PlotAreaMapping(this.WorkbookContext, this.ChartContext));

                    // c:legend
                    var firstLegend = chartFormatsSequence.AxisParentSequences[0].CrtSequences[0].LdSequence;
                    if (firstLegend != null)
                    {
                        firstLegend.Convert(new LegendMapping(this.WorkbookContext, this.ChartContext));
                    }

                    // c:plotVisOnly
                    writeValueElement(Dml.Chart.ElPlotVisOnly, chartFormatsSequence.ShtProps.fPlotVisOnly ? "1" : "0");
                    
                    // c:dispBlanksAs
                    string dispBlanksAs = string.Empty;
                    switch (chartFormatsSequence.ShtProps.mdBlank)
                    {
                        case ShtProps.EmptyCellPlotMode.PlotNothing:
                            dispBlanksAs = "gap";
                            break;

                        case ShtProps.EmptyCellPlotMode.PlotAsZero:
                            dispBlanksAs = "zero";
                            break;

                        case ShtProps.EmptyCellPlotMode.PlotAsInterpolated:
                            dispBlanksAs = "span";
                            break;
                    }
                    if (!string.IsNullOrEmpty(dispBlanksAs))
                    {
                        writeValueElement(Dml.Chart.ElDispBlanksAs, dispBlanksAs);
                    }

                    // c:showDLblsOverMax

                }
                this._writer.WriteEndElement(); // c:chart

                // c:spPr
                if (chartFormatsSequence.FrameSequence != null)
                {
                    chartFormatsSequence.FrameSequence.Convert(new ShapePropertiesMapping(this.WorkbookContext, this.ChartContext));
                }
            }
            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }

        #endregion
    }
}
