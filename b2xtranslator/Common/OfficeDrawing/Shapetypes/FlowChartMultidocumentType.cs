namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(115)]
    public class FlowChartMultidocumentType : ShapeType
    {
        public FlowChartMultidocumentType()
        {
            this.ShapeConcentricFill = false;

            this.Joins = JoinStyle.miter;

            this.Path = "m,20465v810,317,1620,452,2397,725c3077,21325,3790,21417,4405,21597v1620,,2202,-180,2657,-272c7580,21280,8002,21010,8455,20917v422,-135,810,-405,1327,-542c10205,20150,10657,19967,11080,19742v517,-182,970,-407,1425,-590c13087,19017,13605,18745,14255,18610v615,-180,1262,-318,1942,-408c16975,18202,17785,18022,18595,18022r,-1670l19192,16252r808,l20000,14467r722,-75l21597,14392,21597,,2972,r,1815l1532,1815r,1860l,3675,,20465xem1532,3675nfl18595,3675r,12677em2972,1815nfl20000,1815r,12652e";

            this.ConnectorLocations = "10800,0;0,10800;10800,19890;21600,10800";

            this.TextboxRectangle = "0,3675,18595,18022";

        }
    }
}
