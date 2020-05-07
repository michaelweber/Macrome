namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(3)]
    public class OvalType : ShapeType
    {
        public OvalType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.round;

        }
    }
}
