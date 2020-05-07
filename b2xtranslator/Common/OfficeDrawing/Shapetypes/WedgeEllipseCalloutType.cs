using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(63)]
    public class WedgeEllipseCalloutType : ShapeType
    {
        public WedgeEllipseCalloutType()
        {
            this.ShapeConcentricFill = false;

            this.Joins = JoinStyle.miter;

            this.Path = "wr,,21600,21600@15@16@17@18l@21@22xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("val #1");
            this.Formulas.Add("sum 10800 0 #0");
            this.Formulas.Add("sum 10800 0 #1");
            this.Formulas.Add("atan2 @2 @3");
            this.Formulas.Add("sumangle @4 11 0");
            this.Formulas.Add("sumangle @4 0 11");
            this.Formulas.Add("cos 10800 @4");
            this.Formulas.Add("sin 10800 @4");
            this.Formulas.Add("cos 10800 @5");
            this.Formulas.Add("sin 10800 @5");
            this.Formulas.Add("cos 10800 @6");
            this.Formulas.Add("sin 10800 @6");
            this.Formulas.Add("sum 10800 0 @7");
            this.Formulas.Add("sum 10800 0 @8");
            this.Formulas.Add("sum 10800 0 @9");
            this.Formulas.Add("sum 10800 0 @10");
            this.Formulas.Add("sum 10800 0 @11");
            this.Formulas.Add("sum 10800 0 @12");
            this.Formulas.Add("mod @2 @3 0");
            this.Formulas.Add("sum @19 0 10800");
            this.Formulas.Add("if @20 #0 @13");
            this.Formulas.Add("if @20 #1 @14");

            this.AdjustmentValues = "1350,25920";

            this.ConnectorLocations = "10800,0;3163,3163;0,10800;3163,18437;10800,21600;18437,18437;21600,10800;18437,3163;@21,@22";

            this.TextboxRectangle = "3163,3163,18437,18437";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,#1"
            };
            this.Handles.Add(HandleOne);

        }
    }
}
