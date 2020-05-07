namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(114)]
    public class FlowChartDocumentType : ShapeType
    {
        public FlowChartDocumentType()
        {
            this.ShapeConcentricFill = false;

            this.Joins = JoinStyle.miter;

            this.Path = "m,20172v945,400,1887,628,2795,913c3587,21312,4342,21370,5060,21597v2037,,2567,-227,3095,-285c8722,21197,9325,20970,9855,20800v490,-228,945,-400,1472,-740c11817,19887,12347,19660,12875,19375v567,-228,1095,-513,1700,-740c15177,18462,15782,18122,16537,17950v718,-113,1398,-398,2228,-513c19635,17437,20577,17322,21597,17322l21597,,,xe";

            this.ConnectorLocations = "10800,0;0,10800;10800,20400;21600,10800";
            
            this.TextboxRectangle = "0,0,21600,17322";

        }
    }
}
