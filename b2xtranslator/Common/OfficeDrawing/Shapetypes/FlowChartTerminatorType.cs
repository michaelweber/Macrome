namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(116)]
    public class FlowChartTerminatorType : ShapeType
    {
        public FlowChartTerminatorType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.round;

            this.Path = "m3475,qx,10800,3475,21600l18125,21600qx21600,10800,18125,xe";

            this.ConnectorLocations = "Rectangle";

            this.TextboxRectangle = "1018,3163,20582,18437";

        }
    }
}
