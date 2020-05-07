using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(61)]
    public class WedgeRectCalloutType : ShapeType
    {
        public WedgeRectCalloutType()
        {
            this.ShapeConcentricFill = false;

            this.Joins = JoinStyle.miter;

            this.Path = "m,l0@8@12@24,0@9,,21600@6,21600@15@27@7,21600,21600,21600,21600@9@18@30,21600@8,21600,0@7,0@21@33@6,xe";
            
            this.Formulas = new List<string>();
            this.Formulas.Add("sum 10800 0 #0");
            this.Formulas.Add("sum 10800 0 #1");
            this.Formulas.Add("sum #0 0 #1");
            this.Formulas.Add("sum @0 @1 0");
            this.Formulas.Add("sum 21600 0 #0");
            this.Formulas.Add("sum 21600 0 #1");
            this.Formulas.Add("if @0 3600 12600");
            this.Formulas.Add("if @0 9000 18000");
            this.Formulas.Add("if @1 3600 12600");
            this.Formulas.Add("if @1 9000 18000");
            this.Formulas.Add("if @2 0 #0");
            this.Formulas.Add("if @3 @10 0");
            this.Formulas.Add("if #0 0 @11");
            this.Formulas.Add("if @2 @6 #0");
            this.Formulas.Add("if @3 @6 @13");
            this.Formulas.Add("if @5 @6 @14");
            this.Formulas.Add("if @2 #0 21600");
            this.Formulas.Add("if @3 21600 @16");
            this.Formulas.Add("if @4 21600 @17");
            this.Formulas.Add("if @2 #0 @6");
            this.Formulas.Add("if @3 @19 @6");
            this.Formulas.Add("if #1 @6 @20");
            this.Formulas.Add("if @2 @8 #1");
            this.Formulas.Add("if @3 @22 @8");
            this.Formulas.Add("if #0 @8 @23");
            this.Formulas.Add("if @2 21600 #1");
            this.Formulas.Add("if @3 21600 @25");
            this.Formulas.Add("if @5 21600 @26");
            this.Formulas.Add("if @2 #1 @8");
            this.Formulas.Add("if @3 @8 @28");
            this.Formulas.Add("if @4 @8 @29");
            this.Formulas.Add("if @2 #1 0");
            this.Formulas.Add("if @3 @31 0");
            this.Formulas.Add("if #1 0 @32");
            this.Formulas.Add("val #0");
            this.Formulas.Add("val #1");

            this.AdjustmentValues = "1350,25920";

            this.ConnectorLocations = "10800,0;0,10800;10800,21600;21600,10800;@34,@35";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,#1"
            };
            this.Handles.Add(HandleOne);


        }
    }
}
