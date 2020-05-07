using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(65)]
    public class FoldedCornerType : ShapeType
    {
        public FoldedCornerType()
        {
            this.ShapeConcentricFill = false;

            this.Joins = JoinStyle.miter;

            this.Path = "m,l,21600@0,21600,21600@0,21600,xem@0,21600nfl@3@5c@7@9@11@13,21600@0e";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum 21600 0 @0");
            this.Formulas.Add("prod @1 8481 32768");
            this.Formulas.Add("sum @2 @0 0");
            this.Formulas.Add("prod @1 1117 32768");
            this.Formulas.Add("sum @4 @0 0");
            this.Formulas.Add("prod @1 11764 32768");
            this.Formulas.Add("sum @6 @0 0");
            this.Formulas.Add("prod @1 6144 32768");
            this.Formulas.Add("sum @8 @0 0");
            this.Formulas.Add("prod @1 20480 32768");
            this.Formulas.Add("sum @10 @0 0");
            this.Formulas.Add("prod @1 6144 32768");
            this.Formulas.Add("sum @12 @0 0");

            this.AdjustmentValues = "18900";

            this.ConnectorLocations = "Rectangle";

            this.TextboxRectangle = "0,0,21600,@13";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,bottomRight",
                xrange = "10800,21600"
            };
            this.Handles.Add(HandleOne);
            
        }
    }
}
