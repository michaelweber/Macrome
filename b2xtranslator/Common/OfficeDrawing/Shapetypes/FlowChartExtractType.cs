namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(127)]
    class FlowChartExtractType : ShapeType
    {
        public FlowChartExtractType()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m10800,l21600,21600,,21600xe";
            this.ConnectorLocations = "10800,0;5400,10800;10800,21600;16200,10800";
            this.TextboxRectangle = "5400,10800,16200,21600";
        }
    }
}


