using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(83)]
    class QuadArrowCalloutType : ShapeType
    {
        public QuadArrowCalloutType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m@0@0l@3@0@3@2@1@2,10800,0@4@2@5@2@5@0@8@0@8@3@9@3@9@1,21600,10800@9@4@9@5@8@5@8@8@5@8@5@9@4@9,10800,21600@1@9@3@9@3@8@0@8@0@5@2@5@2@4,,10800@2@1@2@3@0@3xe"; 
            this.Formulas = new List<string>();

            this.Formulas.Add("val #0"); 
            this.Formulas.Add("val #1"); 
            this.Formulas.Add("val #2"); 
            this.Formulas.Add("val #3"); 
            this.Formulas.Add("sum 21600 0 #1"); 
            this.Formulas.Add("sum 21600 0 #3"); 
            this.Formulas.Add("sum #0 21600 0"); 
            this.Formulas.Add("prod @6 1 2"); 
            this.Formulas.Add("sum 21600 0 #0"); 
            this.Formulas.Add("sum 21600 0 #2");
            
            this.AdjustmentValues = "5400,8100,2700,9450";
            this.ConnectorLocations = "Rectangle";

            this.TextboxRectangle = "@0,@0,@8,@8";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "topLeft,#0",
                yrange = "@2,@1"
            };
            this.Handles.Add(HandleOne);

            var HandleTwo = new Handle
            {
                position = "#1,topLeft",
                xrange = "@0,@3"
            };
            this.Handles.Add(HandleTwo);

            var HandleThree = new Handle
            {
                position = "#3,#2",
                xrange = "@1,10800",
                yrange = "0,@0"
            };
            this.Handles.Add(HandleThree);


        }
    }
}


