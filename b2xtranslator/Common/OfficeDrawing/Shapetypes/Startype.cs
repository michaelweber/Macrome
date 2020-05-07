namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(12)]
    public class Startype : ShapeType
    {
        public Startype()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;

            this.Path="m10800,l8280,8259,,8259r6720,5146l4200,21600r6600,-5019l17400,21600,14880,13405,21600,8259r-8280,xe";

            this.ConnectorLocations = "10800,0;0,8259;4200,21600;17400,21600;21600,8259";

            this.TextboxRectangle = "6720,8259,14880,15628";

        }
    
    }
}
