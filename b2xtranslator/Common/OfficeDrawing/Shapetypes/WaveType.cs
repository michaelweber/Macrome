using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(64)]
    public class WaveType : ShapeType
    {
        public WaveType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m@28@0c@27@1@26@3@25@0l@21@4c@22@5@23@6@24@4xe";


            this.AdjustmentValues = "2809,10800";
            this.ConnectorLocations = "@35,@0;@38,10800;@37,@4;@36,10800";

            this.ConnectorAngles = "270,180,90,0";

            this.TextboxRectangle = "@31,@33,@32,@34";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0"); 
            this.Formulas.Add("prod @0 41 9"); 
            this.Formulas.Add("prod @0 23 9 ");
            this.Formulas.Add("sum 0 0 @2 ");
            this.Formulas.Add("sum 21600 0 #0"); 
            this.Formulas.Add("sum 21600 0 @1 ");
            this.Formulas.Add("sum 21600 0 @3 ");
            this.Formulas.Add("sum #1 0 10800 ");
            this.Formulas.Add("sum 21600 0 #1 ");
            this.Formulas.Add("prod @8 2 3 ");
            this.Formulas.Add("prod @8 4 3 ");
            this.Formulas.Add("prod @8 2 1 ");
            this.Formulas.Add("sum 21600 0 @9 ");
            this.Formulas.Add("sum 21600 0 @10 ");
            this.Formulas.Add("sum 21600 0 @11 ");
            this.Formulas.Add("prod #1 2 3 ");
            this.Formulas.Add("prod #1 4 3 ");
            this.Formulas.Add("prod #1 2 1 ");
            this.Formulas.Add("sum 21600 0 @15"); 
            this.Formulas.Add("sum 21600 0 @16 ");
            this.Formulas.Add("sum 21600 0 @17 ");
            this.Formulas.Add("if @7 @14 0 ");
            this.Formulas.Add("if @7 @13 @15 ");
            this.Formulas.Add("if @7 @12 @16 ");
            this.Formulas.Add("if @7 21600 @17 ");
            this.Formulas.Add("if @7 0 @20 ");
            this.Formulas.Add("if @7 @9 @19 ");
            this.Formulas.Add("if @7 @10 @18 ");
            this.Formulas.Add("if @7 @11 21600 ");
            this.Formulas.Add("sum @24 0 @21 ");
            this.Formulas.Add("sum @4 0 @0 ");
            this.Formulas.Add("max @21 @25 ");
            this.Formulas.Add("min @24 @28 ");
            this.Formulas.Add("prod @0 2 1 ");
            this.Formulas.Add("sum 21600 0 @33"); 
            this.Formulas.Add("mid @26 @27 ");
            this.Formulas.Add("mid @24 @28 ");
            this.Formulas.Add("mid @22 @23 ");
            this.Formulas.Add("mid @21 @25");

            
            this.Handles = new List<Handle>();
            var handleOne = new Handle
            {
                position = "topLeft,#0",
                yrange = "0,4459"
            };
            this.Handles.Add(handleOne);

            var handleTwo = new Handle
            {
                position = "#1,bottomRight",
                xrange = "8640,12960"
            };
            this.Handles.Add(handleTwo);

        }
    }
}
