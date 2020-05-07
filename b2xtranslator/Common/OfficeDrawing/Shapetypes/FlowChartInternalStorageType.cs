namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(113)]
    public class FlowChartInternalStorageType : ShapeType
    {
        public FlowChartInternalStorageType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m,l,21600r21600,l21600,xem4236,nfl4236,21600em,4236nfl21600,4236e";
            
            this.ConnectorLocations = "Rectangle";

            this.TextboxRectangle = "4236,4236,21600,21600";
        }

    
    }
}
