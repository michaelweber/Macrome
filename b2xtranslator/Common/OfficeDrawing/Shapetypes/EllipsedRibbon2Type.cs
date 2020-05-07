using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(108)]
    public class EllipsedRibbon2Type : ShapeType
    {
        public EllipsedRibbon2Type()
        {
            this.ShapeConcentricFill = false;
            
            this.Joins = JoinStyle.miter;

            this.Path = "wr@9@34@8@35,0@24@0@23@9,0@8@11@0@22@19@22@9@34@8@35@19@23@3@24l@7@36@3@4at@9@31@8@32@3@4@18@30@9@1@8@33@18@28@17@28@9@31@8@32@17@30,0@4l@5@36xear@9@1@8@33@17@28@0@29nfl@17@30ewr@9@1@8@33@18@28@19@29nfl@18@30em@0@23nfl@0@29em@19@23nfl@19@29e";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0"); 
            this.Formulas.Add("val #1");  
            this.Formulas.Add("val #2");  
            this.Formulas.Add("val width");  
            this.Formulas.Add("val height");  
            this.Formulas.Add("prod width 1 8");  
            this.Formulas.Add("prod width 1 2");  
            this.Formulas.Add("prod width 7 8");  
            this.Formulas.Add("prod width 3 2");  
            this.Formulas.Add("sum 0 0 @6");  
            this.Formulas.Add("prod #2 30573 4096");  
            this.Formulas.Add("prod @10 2 1");  
            this.Formulas.Add("sum @10 height #2");  
            this.Formulas.Add("sum @10 #1 0");  
            this.Formulas.Add("prod #1 1 2");  
            this.Formulas.Add("sum @10 @14 0");  
            this.Formulas.Add("sum @12 0 #1");  
            this.Formulas.Add("sum #0 @5 0");  
            this.Formulas.Add("sum width 0 @17");  
            this.Formulas.Add("sum width 0 #0");  
            this.Formulas.Add("sum @6 0 #0");  
            this.Formulas.Add("ellipse @20 width @10");  
            this.Formulas.Add("sum @10 0 @21");  
            this.Formulas.Add("sum @22 @16 @10"); 
            this.Formulas.Add("sum #2 @16 @10");  
            this.Formulas.Add("prod @10 2391 32768");  
            this.Formulas.Add("sum @6 0 @17");  
            this.Formulas.Add("ellipse @26 width @10");  
            this.Formulas.Add("sum @10 #1 @27");  
            this.Formulas.Add("sum @22 #1 0");  
            this.Formulas.Add("sum @12 0 @27");  
            this.Formulas.Add("sum height 0 #2");  
            this.Formulas.Add("sum @10 @12 0");  
            this.Formulas.Add("sum @32 @10 @16");  
            this.Formulas.Add("sum @31 @10 @13");  
            this.Formulas.Add("sum @32 @10 @13");  
            this.Formulas.Add("sum @25 @12 @15");  
            this.Formulas.Add("sum @16 0 @15");  
            this.Formulas.Add("prod @37 2 3");  
            this.Formulas.Add("sum @1 @38 0");  
            this.Formulas.Add("sum #2 @38 0");  
            this.Formulas.Add("max @40 675");  
            this.Formulas.Add("prod width 3 8");  
            this.Formulas.Add("sum @42 0 4"); 

     
            this.AdjustmentValues = "5400,16200,2700"; 

            this.ConnectorLocations = "@6,0;@5,@36;@6,@1;@7,@36";

            this.ConnectorAngles = "270,180,90,0";

            this.TextboxRectangle = "@0,@22,@19,@1";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,topLeft",
                xrange = "@5,@43"
            };
            this.Handles.Add(HandleOne);

            var HandleTwo = new Handle
            {
                position = "center,#1",
                yrange = "@39,@31"
            };
            this.Handles.Add(HandleTwo);

            var HandleThree = new Handle
            {
                position = "topLeft,#2",
                yrange = "@41,@24"
            };
            this.Handles.Add(HandleThree);
        }
    }
}
