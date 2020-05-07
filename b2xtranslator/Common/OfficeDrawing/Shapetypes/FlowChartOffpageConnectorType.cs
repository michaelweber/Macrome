namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(177)]
    class FlowChartOffpageConnectorType : ShapeType
    {
        public FlowChartOffpageConnectorType()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m,l21600,r,17255l10800,21600,,17255xe"; 
            this.ConnectorLocations = "Rectangle";
            this.TextboxRectangle = "0,0,21600,17255";
        }
    }
}


