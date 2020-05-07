using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(156)]
    public class TextWave1 : ShapeType
    {
        public TextWave1()
        {
            this.TextPath = true;
            this.AdjustmentValues = "2809,10800";
            this.Path = "m@25@0c@26@3@27@1@28@0m@21@4c@22@5@23@6@24@4e";
            this.ConnectorLocations = "@35,@0;@38,10800;@37,@4;@36,10800";
            this.ConnectorAngles = "270,180,90,0";

            this.Handles = new List<Handle>();
            var h1 = new Handle
            {
                position = "topLeft,#0",
                yrange = "0,4459"
            };
            var h2 = new Handle
            {
                position = "#1,bottomRight",
                xrange = "8640,12960"
            };
            this.Handles.Add(h1);
            this.Handles.Add(h2);

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("prod @0 41 9");
            this.Formulas.Add("prod @0 23 9");
            this.Formulas.Add("sum 0 0 @2");
            this.Formulas.Add("sum 21600 0 #0");
            this.Formulas.Add("sum 21600 0 @1");
            this.Formulas.Add("sum 21600 0 @3");
            this.Formulas.Add("sum #1 0 10800");
            this.Formulas.Add("sum 21600 0 #1");
            this.Formulas.Add("prod @8 2 3");
            this.Formulas.Add("prod @8 4 3");
            this.Formulas.Add("prod @8 2 1");
            this.Formulas.Add("sum 21600 0 @9");
            this.Formulas.Add("sum 21600 0 @10");
            this.Formulas.Add("sum 21600 0 @11");
            this.Formulas.Add("prod #1 2 3");
            this.Formulas.Add("prod #1 4 3");
            this.Formulas.Add("prod #1 2 1");
            this.Formulas.Add("sum 21600 0 @15");
            this.Formulas.Add("sum 21600 0 @16");
            this.Formulas.Add("sum 21600 0 @17");
            this.Formulas.Add("if @7 @14 0");
            this.Formulas.Add("if @7 @13 @15");
            this.Formulas.Add("if @7 @12 @16");
            this.Formulas.Add("if @7 21600 @17");
            this.Formulas.Add("if @7 0 @20");
            this.Formulas.Add("if @7 @9 @19");
            this.Formulas.Add("if @7 @10 @18");
            this.Formulas.Add("if @7 @11 21600");
            this.Formulas.Add("sum @24 0 @21");
            this.Formulas.Add("sum @4 0 @0");
            this.Formulas.Add("max @21 @25");
            this.Formulas.Add("min @24 @28");
            this.Formulas.Add("prod @0 2 1");
            this.Formulas.Add("sum 21600 0 @33");
            this.Formulas.Add("mid @26 @27");
            this.Formulas.Add("mid @24 @28");
            this.Formulas.Add("mid @22 @23");
            this.Formulas.Add("mid @21 @25");

        }      
    }
}
