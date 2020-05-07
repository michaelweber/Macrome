using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(102)]
    class CurvedRightArrowType : ShapeType
    {
        public CurvedRightArrowType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "ar,0@23@3@22,,0@4,0@15@23@1,0@7@2@13l@2@14@22@8@2@12wa,0@23@3@2@11@26@17,0@15@23@1@26@17@22@15xear,0@23@3,0@4@26@17nfe";
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
            this.Formulas.Add("sum height 0 #2"); 
            this.Formulas.Add("ellipse @9 height @4"); 
            this.Formulas.Add("sum @4 @10 0"); 
            this.Formulas.Add("sum @11 #1 width"); 
            this.Formulas.Add("sum @7 @10 0"); 
            this.Formulas.Add("sum @12 width #0"); 
            this.Formulas.Add("sum @5 0 #0"); 
            this.Formulas.Add("prod @15 1 2"); 
            this.Formulas.Add("mid @4 @7"); 
            this.Formulas.Add("sum #0 #1 width"); 
            this.Formulas.Add("prod @18 1 2"); 
            this.Formulas.Add("sum @17 0 @19"); 
            this.Formulas.Add("val width"); 
            this.Formulas.Add("val height"); 
            this.Formulas.Add("prod height 2 1"); 
            this.Formulas.Add("sum @17 0 @4"); 
            this.Formulas.Add("ellipse @24 @4 height"); 
            this.Formulas.Add("sum height 0 @25"); 
            this.Formulas.Add("sum @8 128 0"); 
            this.Formulas.Add("prod @5 1 2"); 
            this.Formulas.Add("sum @5 0 128"); 
            this.Formulas.Add("sum #0 @17 @12"); 
            this.Formulas.Add("ellipse @20 @4 height"); 
            this.Formulas.Add("sum width 0 #0"); 
            this.Formulas.Add("prod @32 1 2"); 
            this.Formulas.Add("prod height height 1"); 
            this.Formulas.Add("prod @9 @9 1"); 
            this.Formulas.Add("sum @34 0 @35"); 
            this.Formulas.Add("sqrt @36"); 
            this.Formulas.Add("sum @37 height 0"); 
            this.Formulas.Add("prod width height @38"); 
            this.Formulas.Add("sum @39 64 0"); 
            this.Formulas.Add("prod #0 1 2"); 
            this.Formulas.Add("ellipse @33 @41 height"); 
            this.Formulas.Add("sum height 0 @42"); 
            this.Formulas.Add("sum @43 64 0"); 
            this.Formulas.Add("prod @4 1 2"); 
            this.Formulas.Add("sum #1 0 @45"); 
            this.Formulas.Add("prod height 4390 32768");
            this.Formulas.Add("prod height 28378 32768");

            this.AdjustmentValues = "12960,19440,14400";
            this.ConnectorLocations = "0,@17;@2,@14;@22,@8;@2,@12;@22,@16";
            this.ConnectorAngles = "180,90,0,0,0";

            this.TextboxRectangle = "@47,@45,@48,@46";
           
            this.Handles = new List<Handle>();

            var HandleOne = new Handle
            {
                position = "bottomRight,#0",
                yrange = "@40,@29"
            };
            this.Handles.Add(HandleOne);
            
            var HandleTwo = new Handle();
            HandleOne.position="bottomRight,#1"; 
            HandleOne.yrange="@27,@21";  
            this.Handles.Add(HandleTwo);

            var HandleThree = new Handle
            {
                position = "#2,bottomRight",
                xrange = "@44,@22"
            };
            this.Handles.Add(HandleThree);
        }
    }
}
