namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(135)]
    class FlowChartDelayType : ShapeType
    {
        public FlowChartDelayType()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m10800,qx21600,10800,10800,21600l,21600,,xe";
            this.ConnectorLocations = "Rectangle";
            this.TextboxRectangle = "0,3163,18437,18437";
        }
    }
}


