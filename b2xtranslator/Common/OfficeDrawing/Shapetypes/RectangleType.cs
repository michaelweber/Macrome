namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(1)]
    public class RectangleType : ShapeType
    {
        public RectangleType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m,l,21600r21600,l21600,xe";
        }
    }
}
