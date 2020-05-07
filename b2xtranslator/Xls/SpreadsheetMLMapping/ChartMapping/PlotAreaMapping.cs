

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class PlotAreaMapping : AbstractChartMapping,
        IMapping<ChartFormatsSequence>
    {

        public PlotAreaMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        #region IMapping<ChartFormatsSequence> Members

        /// <summary>
        /// <complexType name="CT_PlotArea">
        ///     <sequence>
        ///         <element name="layout" type="CT_Layout" minOccurs="0" maxOccurs="1"/>
        ///         <choice minOccurs="1" maxOccurs="unbounded">
        ///             <element name="areaChart" type="CT_AreaChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="area3DChart" type="CT_Area3DChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="lineChart" type="CT_LineChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="line3DChart" type="CT_Line3DChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="stockChart" type="CT_StockChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="radarChart" type="CT_RadarChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="scatterChart" type="CT_ScatterChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="pieChart" type="CT_PieChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="pie3DChart" type="CT_Pie3DChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="doughnutChart" type="CT_DoughnutChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="barChart" type="CT_BarChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="bar3DChart" type="CT_Bar3DChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="ofPieChart" type="CT_OfPieChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="surfaceChart" type="CT_SurfaceChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="surface3DChart" type="CT_Surface3DChart" minOccurs="1" maxOccurs="1"/>
        ///             <element name="bubbleChart" type="CT_BubbleChart" minOccurs="1" maxOccurs="1"/>
        ///         </choice>
        ///         <choice minOccurs="0" maxOccurs="unbounded">
        ///             <element name="valAx" type="CT_ValAx" minOccurs="1" maxOccurs="1"/>
        ///             <element name="catAx" type="CT_CatAx" minOccurs="1" maxOccurs="1"/>
        ///             <element name="dateAx" type="CT_DateAx" minOccurs="1" maxOccurs="1"/>
        ///             <element name="serAx" type="CT_SerAx" minOccurs="1" maxOccurs="1"/>
        ///         </choice>
        ///         <element name="dTable" type="CT_DTable" minOccurs="0" maxOccurs="1"/>
        ///         <element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1"/>
        ///         <element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
        ///     </sequence>
        /// </complexType>
        /// </summary>
        public void Apply(ChartFormatsSequence chartFormatsSequence)
        {
            // c:plotArea
            this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElPlotArea, Dml.Chart.Ns);
            {
                // c:layout
                if (chartFormatsSequence.ShtProps.fManPlotArea && chartFormatsSequence.CrtLayout12A != null)
                {
                    chartFormatsSequence.CrtLayout12A.Convert(new LayoutMapping(this.WorkbookContext, this.ChartContext));
                }

                // chart groups
                foreach (var axisParentSequence in chartFormatsSequence.AxisParentSequences)
                {
                    foreach (var crtSequence in axisParentSequence.CrtSequences)
                    {
                        // The Chart3d record specifies that the plot area, axis group, and chart group are rendered 
                        // in a 3-D scene, rather than a 2-D scene, and specifies properties of the 3-D scene. If this 
                        // record exists in the chart sheet substream, the chart sheet substream MUST have exactly one 
                        // chart group. This record MUST NOT exist in a bar of pie, bubble, doughnut,
                        // filled radar, pie of pie, radar, or scatter chart group.
                        //
                        bool is3DChart = (crtSequence.Chart3d != null);

                        // area chart
                        if (crtSequence.ChartType is Area)
                        {
                            crtSequence.Convert(new AreaChartMapping(this.WorkbookContext, this.ChartContext, is3DChart));
                        }
                        // bar and column chart
                        else if (crtSequence.ChartType is Bar)
                        {
                            crtSequence.Convert(new BarChartMapping(this.WorkbookContext, this.ChartContext, is3DChart));
                        }
                        // OfPieChart (Bar of pie / Pie of Pie)
                        else if (crtSequence.ChartType is BopPop)
                        {
                            crtSequence.Convert(new OfPieChartMapping(this.WorkbookContext, this.ChartContext, is3DChart));
                        }
                        // bubbleChart
                        else if (crtSequence.ChartType is Scatter && ((Scatter)crtSequence.ChartType).fBubbles)
                        {
                            crtSequence.Convert(new BubbleChartMapping(this.WorkbookContext, this.ChartContext, is3DChart));
                        }
                        // scatterChart
                        else if (crtSequence.ChartType is Scatter && !((Scatter)crtSequence.ChartType).fBubbles)
                        {
                            crtSequence.Convert(new ScatterChartMapping(this.WorkbookContext, this.ChartContext, is3DChart));
                        }
                        // lineChart and stockChart
                        else if (crtSequence.ChartType is Line)
                        {
                            crtSequence.Convert(new LineChartMapping(this.WorkbookContext, this.ChartContext, is3DChart));
                        }
                        // doughnutChart and pieChart (they differ by ((Pie)crtSequence.ChartType).pcDonut
                        else if (crtSequence.ChartType is Pie)
                        {
                            crtSequence.Convert(new PieChartMapping(this.WorkbookContext, this.ChartContext, is3DChart));
                        }
                        // radarChart
                        else if (crtSequence.ChartType is Radar)
                        {
                            // RadarArea (or "Filled Radar") has the radarStyle set to "filled")
                            crtSequence.Convert(new RadarChartMapping(this.WorkbookContext, this.ChartContext, is3DChart));
                        }
                        // surfaceChart
                        else if (crtSequence.ChartType is Surf)
                        {
                            crtSequence.Convert(new SurfaceChartMapping(this.WorkbookContext, this.ChartContext, is3DChart));
                        }
                    }
                }

                // axis groups
                foreach (var axisParentSequence in chartFormatsSequence.AxisParentSequences)
                {
                    // NOTE: AxisParent.iax must be 0 for the primary axis group
                    var axesSequence = axisParentSequence.AxesSequence;
                    if (axesSequence != null)
                    {
                        axesSequence.Convert(new AxisMapping(this.WorkbookContext, this.ChartContext));
                    }
                }
                // c:spPr
                if (chartFormatsSequence.AxisParentSequences.Count > 0
                    && chartFormatsSequence.AxisParentSequences[0].AxesSequence != null
                    && chartFormatsSequence.AxisParentSequences[0].AxesSequence.PlotArea != null)
                {
                    chartFormatsSequence.AxisParentSequences[0].AxesSequence.Frame.Convert(new ShapePropertiesMapping(this.WorkbookContext, this.ChartContext));
                }
            }
            this._writer.WriteEndElement(); // c:plotArea
        }
        #endregion
    }
}
