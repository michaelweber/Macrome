using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(99)]
    class CircularArrowType : ShapeType
    {
        public CircularArrowType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "al10800,10800@8@8@4@6,10800,10800,10800,10800@9@7l@30@31@17@18@24@25@15@16@32@33xe"; 
            this.Formulas = new List<string>();

            this.Formulas.Add("val #1"); 
            this.Formulas.Add("val #0"); 
            this.Formulas.Add("sum #1 0 #0"); 
            this.Formulas.Add("val 10800"); 
            this.Formulas.Add("sum 0 0 #1"); 
            this.Formulas.Add("sumangle @2 360 0"); 
            this.Formulas.Add("if @2 @2 @5"); 
            this.Formulas.Add("sum 0 0 @6"); 
            this.Formulas.Add("val #2"); 
            this.Formulas.Add("sum 0 0 #0");
            this.Formulas.Add("sum #2 0 2700"); 
            this.Formulas.Add("cos @10 #1 ");
            this.Formulas.Add("sin @10 #1 ");
            this.Formulas.Add("cos 13500 #1"); 
            this.Formulas.Add("sin 13500 #1 ");
            this.Formulas.Add("sum @11 10800 0"); 
            this.Formulas.Add("sum @12 10800 0 ");
            this.Formulas.Add("sum @13 10800 0 ");
            this.Formulas.Add("sum @14 10800 0 ");
            this.Formulas.Add("prod #2 1 2 ");
            this.Formulas.Add("sum @19 5400 0"); 
            this.Formulas.Add("cos @20 #1"); 
            this.Formulas.Add("sin @20 #1"); 
            this.Formulas.Add("sum @21 10800 0 ");
            this.Formulas.Add("sum @12 @23 @22"); 
            this.Formulas.Add("sum @22 @23 @11"); 
            this.Formulas.Add("cos 10800 #1"); 
            this.Formulas.Add("sin 10800 #1"); 
            this.Formulas.Add("cos #2 #1 ");
            this.Formulas.Add("sin #2 #1 ");
            this.Formulas.Add("sum @26 10800 0"); 
            this.Formulas.Add("sum @27 10800 0"); 
            this.Formulas.Add("sum @28 10800 0"); 
            this.Formulas.Add("sum @29 10800 0"); 
            this.Formulas.Add("sum @19 5400 0 ");
            this.Formulas.Add("cos @34 #0 ");
            this.Formulas.Add("sin @34 #0 ");
            this.Formulas.Add("mid #0 #1 ");
            this.Formulas.Add("sumangle @37 180 0 ");
            this.Formulas.Add("if @2 @37 @38"); 
            this.Formulas.Add("cos 10800 @39 ");
            this.Formulas.Add("sin 10800 @39 ");
            this.Formulas.Add("cos #2 @39 ");
            this.Formulas.Add("sin #2 @39 ");
            this.Formulas.Add("sum @40 10800 0"); 
            this.Formulas.Add("sum @41 10800 0 ");
            this.Formulas.Add("sum @42 10800 0 ");
            this.Formulas.Add("sum @43 10800 0 ");
            this.Formulas.Add("sum @35 10800 0 ");
            this.Formulas.Add("sum @36 10800 0");

    
            this.AdjustmentValues = "-11796480,,5400";
            this.ConnectorLocations = "@44,@45;@48,@49;@46,@47;@17,@18;@24,@25;@15,@16";

            this.TextboxRectangle = "3163,3163,18437,18437";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "@3,#0",
                polar = "10800,10800"
            };
            this.Handles.Add(HandleOne);

            var HandleTwo = new Handle
            {
                position = "#2,#1",
                polar = "10800,10800",
                radiusrange = "0,10800"
            };
            this.Handles.Add(HandleTwo);

        }
    }
}


