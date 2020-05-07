using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(67)]
    public class DownArrowType :ShapeType
    {
        public DownArrowType()
        {
            this.ShapeConcentricFill = false;

            this.Joins = JoinStyle.miter;


            this.Path = "m0@0l@1@0@1,0@2,0@2@0,21600@0,10800,21600xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("val #1");
            this.Formulas.Add("sum height 0 #1");
            this.Formulas.Add("sum 10800 0 #1");
            this.Formulas.Add("sum width 0 #0");
            this.Formulas.Add("prod @4 @3 10800");
            this.Formulas.Add("sum width 0 @5");

            this.AdjustmentValues = "16200,5400";

            this.ConnectorLocations = "10800,0;0,@0;10800,21600;21600,@0";

            this.ConnectorAngles = "270,180,90,0";
            this.TextboxRectangle = "@1,0,@2,@6";

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
