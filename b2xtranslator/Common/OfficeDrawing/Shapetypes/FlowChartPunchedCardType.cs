namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(121)]
    class FlowChartPunchedCardType : ShapeType
    {
        public FlowChartPunchedCardType()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m4321,l21600,r,21600l,21600,,4338xe"; 
            this.ConnectorLocations = "Rectangle";
            this.TextboxRectangle = "0,4321,21600,21600";
        }
    }
}


