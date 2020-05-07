using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(68)]
    public class UpArrowType :ShapeType
    {
        public UpArrowType()
        {
            this.ShapeConcentricFill = false;

            this.Joins = JoinStyle.miter;
            this.Path = "m0@0l@1@0@1,21600@2,21600@2@0,21600@0,10800,xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("val #1");
            this.Formulas.Add("sum 21600 0 #1");
            this.Formulas.Add("prod #0 #1 10800");
            this.Formulas.Add("sum #0 0 @3");

            this.AdjustmentValues = "5400,5400";

            this.ConnectorLocations = "10800,0;0,@0;10800,21600;21600,@0";

            this.ConnectorAngles = "270,180,90,0";

            this.TextboxRectangle = "@1,@4,@2,21600";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#1,#0",
                xrange = "0,10800",
                yrange = "0,21600"
            };
            this.Handles.Add(HandleOne);
        }
    }
}
