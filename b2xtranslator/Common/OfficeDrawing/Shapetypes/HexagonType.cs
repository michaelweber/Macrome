using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(9)]
    public class HexagonType : ShapeType
    {
        public HexagonType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m@0,l,10800@0,21600@1,21600,21600,10800@1,xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum width 0 #0");
            this.Formulas.Add("sum height 0 #0");
            this.Formulas.Add("prod @0 2929 10000");
            this.Formulas.Add("sum width 0 @3");
            this.Formulas.Add("sum height 0 @3");

            this.AdjustmentValues = "5400";
            
            this.ConnectorLocations = "Rectangle";

            this.TextboxRectangle = "1800,1800,19800,19800;3600,3600,18000,18000;6300,6300,15300,15300";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,topLeft",
                xrange = "0,10800"
            };
            this.Handles.Add(HandleOne);

        }
    }
}
