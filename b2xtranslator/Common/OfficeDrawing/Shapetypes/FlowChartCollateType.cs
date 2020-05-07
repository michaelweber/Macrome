namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(125)]
    class FlowChartCollateType : ShapeType
    {
        public FlowChartCollateType()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m21600,21600l,21600,21600,,,xe";
            this.ConnectorLocations = "10800,0;10800,10800;10800,21600";
            this.TextboxRectangle = "5400,5400,16200,16200";

        }
    }
}


