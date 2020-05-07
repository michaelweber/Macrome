using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(7)]
    public class ParallelogramType : ShapeType
    {
        public ParallelogramType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m@0,l,21600@1,21600,21600,xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum width 0 #0");
            this.Formulas.Add("prod #0 1 2");
            this.Formulas.Add("sum width 0 @2");
            this.Formulas.Add("mid #0 width");
            this.Formulas.Add("mid @1 0");
            this.Formulas.Add("prod height width #0");
            this.Formulas.Add("prod @6 1 2");
            this.Formulas.Add("sum height 0 @7");
            this.Formulas.Add("prod width 1 2");
            this.Formulas.Add("sum #0 0 @9");
            this.Formulas.Add("if @10 @8 0");
            this.Formulas.Add("if @10 @7 height");

            this.AdjustmentValues = "5400";
            
            this.ConnectorLocations = "@4,0;10800,@11;@3,10800;@5,21600;10800,@12;@2,10800";

            this.TextboxRectangle = "1800,1800,19800,19800;8100,8100,13500,13500;10800,10800,10800,10800";

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
