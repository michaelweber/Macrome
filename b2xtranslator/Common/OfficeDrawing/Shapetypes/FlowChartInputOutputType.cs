namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(111)]
    public class FlowChartInputOutputType :ShapeType
    {
        public FlowChartInputOutputType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m4321,l21600,,17204,21600,,21600xe";

            this.ConnectorLocations = "2961,0;10800,0;2161,10800;8602,21600;10800,21600;19402,10800";

            this.TextboxRectangle = "4321,0,17204,21600";

        }
    }
}
