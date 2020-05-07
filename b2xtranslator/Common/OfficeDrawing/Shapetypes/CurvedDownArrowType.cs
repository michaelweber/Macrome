using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(105)]
    class CurvedDownArrowType : ShapeType
    {
        public CurvedDownArrowType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "wr,0@3@23,0@22@4,0@15,0@1@23@7,0@13@2l@14@2@8@22@12@2at,0@3@23@11@2@17@26@15,0@1@23@17@26@15@22xewr,0@3@23@4,0@17@26nfe";
            this.Formulas = new List<string>();

            this.Formulas.Add("val #0"); 
            this.Formulas.Add("val #1"); 
            this.Formulas.Add("val #2 ");
            this.Formulas.Add("sum #0 width #1 ");
            this.Formulas.Add("prod @3 1 2 ");
            this.Formulas.Add("sum #1 #1 width"); 
            this.Formulas.Add("sum @5 #1 #0 ");
            this.Formulas.Add("prod @6 1 2");             
            this.Formulas.Add("mid width #0 ");
            this.Formulas.Add("sum height 0 #2 ");
            this.Formulas.Add("ellipse @9 height @4"); 
            this.Formulas.Add("sum @4 @10 0"); 
            this.Formulas.Add("sum @11 #1 width"); 
            this.Formulas.Add("sum @7 @10 0"); 
            this.Formulas.Add("sum @12 width #0 ");
            this.Formulas.Add("sum @5 0 #0 ");
            this.Formulas.Add("prod @15 1 2"); 
            this.Formulas.Add("mid @4 @7 ");
            this.Formulas.Add("sum #0 #1 width"); 
            this.Formulas.Add("prod @18 1 2 ");
            this.Formulas.Add("sum @17 0 @19 ");
            this.Formulas.Add("val width ");
            this.Formulas.Add("val height ");
            this.Formulas.Add("prod height 2 1"); 
            this.Formulas.Add("sum @17 0 @4 ");
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
            this.ConnectorLocations = "@17,0;@16,@22;@12,@2;@8,@22;@14,@2";
            this.ConnectorAngles = "270,90,90,90,0";

            this.TextboxRectangle = "@45,@47,@46,@48";

            this.Handles = new List<Handle>();


            var HandleOne = new Handle
            {
                position = "#0,bottomRight",
                xrange = "@40,@29"
            };
            this.Handles.Add(HandleOne);

            var HandleTwo = new Handle();
            HandleOne.position="#1,bottomRight";
            HandleOne.xrange="@27,@21";
            this.Handles.Add(HandleTwo);

            var HandleThree = new Handle
            {
                position = "bottomRight,#2",
                yrange = "@44,@22"
            };
            this.Handles.Add(HandleThree);
        }
    }
}
