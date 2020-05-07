using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(103)]
    class CurvedLeftArrowType : ShapeType
    {
        public CurvedLeftArrowType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "wr@22,0@21@3,,0@21@4@22@14@21@1@21@7@2@12l@2@13,0@8@2@11at@22,0@21@3@2@10@24@16@22@14@21@1@24@16,0@14xear@22@14@21@1@21@7@24@16nfe";
            this.Formulas = new List<string>();


            this.Formulas.Add("val #0"); 
            this.Formulas.Add("val #1"); 
            this.Formulas.Add("val #2"); 
            this.Formulas.Add("sum #0 width #1"); 
            this.Formulas.Add("prod @3 1 2"); 
            this.Formulas.Add("sum #1 #1 width"); 
            this.Formulas.Add("sum @5 #1 #0"); 
            this.Formulas.Add("prod @6 1 2"); 
            this.Formulas.Add("mid width #0"); 
            this.Formulas.Add("ellipse #2 height @4"); 
            this.Formulas.Add("sum @4 @9 0"); 
            this.Formulas.Add("sum @10 #1 width"); 
            this.Formulas.Add("sum @7 @9 0"); 
            this.Formulas.Add("sum @11 width #0"); 
            this.Formulas.Add("sum @5 0 #0"); 
            this.Formulas.Add("prod @14 1 2"); 
            this.Formulas.Add("mid @4 @7"); 
            this.Formulas.Add("sum #0 #1 width"); 
            this.Formulas.Add("prod @17 1 2"); 
            this.Formulas.Add("sum @16 0 @18"); 
            this.Formulas.Add("val width"); 
            this.Formulas.Add("val height"); 
            this.Formulas.Add("sum 0 0 height"); 
            this.Formulas.Add("sum @16 0 @4"); 
            this.Formulas.Add("ellipse @23 @4 height"); 
            this.Formulas.Add("sum @8 128 0"); 
            this.Formulas.Add("prod @5 1 2"); 
            this.Formulas.Add("sum @5 0 128"); 
            this.Formulas.Add("sum #0 @16 @11"); 
            this.Formulas.Add("sum width 0 #0"); 
            this.Formulas.Add("prod @29 1 2"); 
            this.Formulas.Add("prod height height 1"); 
            this.Formulas.Add("prod #2 #2 1"); 
            this.Formulas.Add("sum @31 0 @32"); 
            this.Formulas.Add("sqrt @33"); 
            this.Formulas.Add("sum @34 height 0"); 
            this.Formulas.Add("prod width height @35"); 
            this.Formulas.Add("sum @36 64 0"); 
            this.Formulas.Add("prod #0 1 2"); 
            this.Formulas.Add("ellipse @30 @38 height"); 
            this.Formulas.Add("sum @39 0 64"); 
            this.Formulas.Add("prod @4 1 2"); 
            this.Formulas.Add("sum #1 0 @41"); 
            this.Formulas.Add("prod height 4390 32768");
            this.Formulas.Add("prod height 28378 32768");


            this.AdjustmentValues = "12960,19440,7200";
            this.ConnectorLocations = "0,@15;@2,@11;0,@8;@2,@13;@21,@16";
            this.ConnectorAngles = "180,180,180,90,0";

            this.TextboxRectangle = "@43,@41,@44,@42";

            this.Handles = new List<Handle>();

            var HandleOne = new Handle
            {
                position = "topLeft,#0",
                yrange = "@37,@27"
            };
            this.Handles.Add(HandleOne);

            var HandleTwo = new Handle();
            HandleOne.position="topLeft,#1";
            HandleOne.yrange="@25,@20";
            this.Handles.Add(HandleTwo);

            var HandleThree = new Handle
            {
                position = "#2,bottomRight",
                xrange = "0,@40"
            };
            this.Handles.Add(HandleThree);
        }
    }
}
