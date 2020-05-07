using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(16)]
    public class CubeType :ShapeType
    {
        public CubeType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;


            this.Path = "m@0,l0@0,,21600@1,21600,21600@2,21600,xem0@0nfl@1@0,21600,em@1@0nfl@1,21600e";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum width 0 #0");
            this.Formulas.Add("sum height 0 #0");
            this.Formulas.Add("mid height #0");
            this.Formulas.Add("prod @1 1 2");
            this.Formulas.Add("prod @2 1 2");
            this.Formulas.Add("mid width #0");

            this.AdjustmentValues = "5400";

            this.ConnectorLocations = "@6,0;@4,@0;0,@3;@4,21600;@1,@3;21600,@5";

            this.ConnectorAngles = "270,270,180,90,0,0";

            this.TextboxRectangle = "0,@0,@1,21600";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "topLeft,#0",
                switchHandle = "true",
                yrange = "0,21600"
            };
            this.Handles.Add(HandleOne);

            this.Limo = "10800,10800";
        }
    }
}
