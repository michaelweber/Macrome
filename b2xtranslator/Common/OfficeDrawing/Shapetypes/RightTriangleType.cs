namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(6)]
    public class RightTriangleType : ShapeType
    {
        public RightTriangleType()
        {

            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m,l,21600r21600,xe";

            this.ConnectorLocations = "0,0;0,10800;0,21600;10800,21600;21600,21600;10800,10800";

            this.TextboxRectangle = "1800,12600,12600,19800";

        }
    }
}
