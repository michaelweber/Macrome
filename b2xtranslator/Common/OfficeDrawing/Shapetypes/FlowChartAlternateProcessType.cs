using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(176)]
    public class FlowChartAlternateProcessType : ShapeType
    {
        public FlowChartAlternateProcessType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m@0,qx0@0l0@2qy@0,21600l@1,21600qx21600@2l21600@0qy@1,xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum width 0 #0");
            this.Formulas.Add("sum height 0 #0");
            this.Formulas.Add("prod @0 2929 10000");
            this.Formulas.Add("sum width 0 @3");
            this.Formulas.Add("sum height 0 @3");
            this.Formulas.Add("val width");
            this.Formulas.Add("val height");
            this.Formulas.Add("prod width 1 2");
            this.Formulas.Add("prod height 1 2");

            this.AdjustmentValues = "2700";
            
            this.ConnectorLocations = "@8,0;0,@9;@8,@7;@6,@9";

            this.TextboxRectangle = "@3,@3,@4,@5";

            this.Limo = "10800,10800";

        }
    }
}
