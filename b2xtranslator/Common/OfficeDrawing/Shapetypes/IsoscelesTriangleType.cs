using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(5)]
    public class IsoscelesTriangleType : ShapeType
    {
        public IsoscelesTriangleType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;

            this.Path = "m@0,l,21600r21600,xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("prod #0 1 2");
            this.Formulas.Add("sum @1 10800 0");

            this.AdjustmentValues = "10800";

            this.ConnectorLocations = "@0,0;@1,10800;0,21600;10800,21600;21600,21600;@2,10800";

            this.TextboxRectangle = "0,10800,10800,18000;5400,10800,16200,18000;10800,10800,21600,18000;0,7200,7200,21600;7200,7200,14400,21600;14400,7200,21600,21600";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,topLeft",
                xrange = "0,21600"
            };
            this.Handles.Add(HandleOne);
        }
    }
}
