using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(89)]
    class LeftUpArrowType : ShapeType
    {
        public LeftUpArrowType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m@4,l@0@2@5@2@5@5@2@5@2@0,0@4@2,21600@2@1@1@1@1@2,21600@2xe";
            this.Formulas = new List<string>();

            this.Formulas.Add("val #0"); 
            this.Formulas.Add("val #1");  
            this.Formulas.Add("val #2");  
            this.Formulas.Add("prod #0 1 2");  
            this.Formulas.Add("sum @3 10800 0");  
            this.Formulas.Add("sum 21600 #0 #1");  
            this.Formulas.Add("sum #1 #2 0");  
            this.Formulas.Add("prod @6 1 2");  
            this.Formulas.Add("prod #1 2 1");  
            this.Formulas.Add("sum @8 0 21600");  
            this.Formulas.Add("sum @5 0 @4");  
            this.Formulas.Add("sum #0 0 @4");
            this.Formulas.Add("prod @2 @10 @11"); 

            this.AdjustmentValues = "9257,18514,6171";
            this.ConnectorLocations = "@4,0;@0,@2;@2,@0;0,@4;@2,21600;@7,@1;@1,@7;21600,@2";
            this.ConnectorAngles = "270,180,270,180,90,90,0,0";

            this.TextboxRectangle = "@12,@5,@1,@1;@5,@12,@1,@1";
           
            this.Handles = new List<Handle>();

            var HandleOne = new Handle
            {
                position = "#0,topLeft",
                xrange = "@2,@9"
            };
            this.Handles.Add(HandleOne);

            var HandleTwo = new Handle
            {
                position = "#1,#2",
                xrange = "@4,21600",
                yrange = "0,@0"
            };
            this.Handles.Add(HandleTwo);

        }
    }
}
