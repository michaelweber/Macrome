using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(107)]
    public class EllipseRibbonType : ShapeType
    {
        public EllipseRibbonType()
        {
            this.ShapeConcentricFill = false;
            
            this.Joins = JoinStyle.miter;

            this.Path = "ar@9@38@8@37,0@27@0@26@9@13@8@4@0@25@22@25@9@38@8@37@22@26@3@27l@7@40@3,wa@9@35@8@10@3,0@21@33@9@36@8@1@21@31@20@31@9@35@8@10@20@33,,l@5@40xewr@9@36@8@1@20@31@0@32nfl@20@33ear@9@36@8@1@21@31@22@32nfl@21@33em@0@26nfl@0@32em@22@26nfl@22@32e";

            this.Formulas = new List<string>();


            
            this.Formulas.Add("val #0"); 
            this.Formulas.Add("val #1"); 
            this.Formulas.Add("val #2 ");
            this.Formulas.Add("val width ");
            this.Formulas.Add("val height ");
            this.Formulas.Add("prod width 1 8"); 
            this.Formulas.Add("prod width 1 2 ");
            this.Formulas.Add("prod width 7 8 ");
            this.Formulas.Add("prod width 3 2 ");
            this.Formulas.Add("sum 0 0 @6 ");
            this.Formulas.Add("sum height 0 #2"); 
            this.Formulas.Add("prod @10 30573 4096"); 
            this.Formulas.Add("prod @11 2 1 ");
            this.Formulas.Add("sum height 0 @12"); 
            this.Formulas.Add("sum @11 #2 0 ");
            this.Formulas.Add("sum @11 height #1"); 
            this.Formulas.Add("sum height 0 #1 ");
            this.Formulas.Add("prod @16 1 2 ");
            this.Formulas.Add("sum @11 @17 0 ");
            this.Formulas.Add("sum @14 #1 height"); 
            this.Formulas.Add("sum #0 @5 0 ");
            this.Formulas.Add("sum width 0 @20"); 
            this.Formulas.Add("sum width 0 #0"); 
            this.Formulas.Add("sum @6 0 #0"); 
            this.Formulas.Add("ellipse @23 width @11 ");
            this.Formulas.Add("sum @24 height @11 ");
            this.Formulas.Add("sum @25 @11 @19 ");
            this.Formulas.Add("sum #2 @11 @19 ");
            this.Formulas.Add("prod @11 2391 32768 ");
            this.Formulas.Add("sum @6 0 @20 ");
            this.Formulas.Add("ellipse @29 width @11 ");
            this.Formulas.Add("sum #1 @30 @11 ");
            this.Formulas.Add("sum @25 #1 height ");
            this.Formulas.Add("sum height @30 @14 ");
            this.Formulas.Add("sum @11 @14 0 ");
            this.Formulas.Add("sum height 0 @34 ");
            this.Formulas.Add("sum @35 @19 @11 ");
            this.Formulas.Add("sum @10 @15 @11 ");
            this.Formulas.Add("sum @35 @15 @11 ");
            this.Formulas.Add("sum @28 @14 @18 ");
            this.Formulas.Add("sum height 0 @39 ");
            this.Formulas.Add("sum @19 0 @18 ");
            this.Formulas.Add("prod @41 2 3 ");
            this.Formulas.Add("sum #1 0 @42 ");
            this.Formulas.Add("sum #2 0 @42 ");
            this.Formulas.Add("min @44 20925 ");
            this.Formulas.Add("prod width 3 8 ");
            this.Formulas.Add("sum @46 0 4");
              
            this.AdjustmentValues = "5400,5400,18900"; 

            this.ConnectorLocations = "@6,@1;@5,@40;@6,@4;@7,@40";

            this.ConnectorAngles = "270,180,90,0";

            this.TextboxRectangle = "@0,@1,@22,@25";
                
            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,bottomRight",
                xrange = "@5,@47"
            };
            this.Handles.Add(HandleOne);

            var HandleTwo = new Handle
            {
                position = "center,#1",
                yrange = "@10,@43"
            };
            this.Handles.Add(HandleTwo);

            var HandleThree = new Handle
            {
                position = "topLeft,#2",
                yrange = "@27,@45"
            };
            this.Handles.Add(HandleThree);
        }
    }
}
