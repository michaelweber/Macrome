namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(118)]
    public class FlowChartManualInputType :ShapeType
    {
        public FlowChartManualInputType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m,4292l21600,r,21600l,21600xe";

            this.ConnectorLocations = "10800,2146;0,10800;10800,21600;21600,10800";

            this.TextboxRectangle = "0,4291,21600,21600";

        }
    }
}
