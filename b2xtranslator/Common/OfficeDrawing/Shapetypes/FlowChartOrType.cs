namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(124)]
    class FlowChartOrType : ShapeType
    {
        public FlowChartOrType()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m10800,qx,10800,10800,21600,21600,10800,10800,xem,10800nfl21600,10800em10800,nfl10800,21600e";
            this.ConnectorLocations = "10800,0;3163,3163;0,10800;3163,18437;10800,21600;18437,18437;21600,10800;18437,3163";
            this.TextboxRectangle = "3163,3163,18437,18437";
        }
    }
}


