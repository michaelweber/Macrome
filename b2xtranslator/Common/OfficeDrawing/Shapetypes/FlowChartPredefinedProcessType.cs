namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(112)]
    public class FlowChartPredefinedProcessType :ShapeType
    {
        public FlowChartPredefinedProcessType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m,l,21600r21600,l21600,xem2610,nfl2610,21600em18990,nfl18990,21600e";

            this.ConnectorLocations = "Rectangle";

            this.TextboxRectangle = "2610,0,18990,21600";

        }
    }
}
