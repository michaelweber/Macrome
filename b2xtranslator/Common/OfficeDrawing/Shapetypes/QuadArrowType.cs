using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(76)]
    class QuadArrowType : ShapeType
    {
        public QuadArrowType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m10800,l@0@2@1@2@1@1@2@1@2@0,,10800@2@3@2@4@1@4@1@5@0@5,10800,21600@3@5@4@5@4@4@5@4@5@3,21600,10800@5@0@5@1@4@1@4@2@3@2xe";
            this.Formulas = new List<string>();
            this.Formulas.Add("val #0 ");
            this.Formulas.Add("val #1");
            this.Formulas.Add("val #2 ");
            this.Formulas.Add("sum 21600 0 #0 ");
            this.Formulas.Add("sum 21600 0 #1 ");
            this.Formulas.Add("sum 21600 0 #2 ");
            this.Formulas.Add("sum #0 0 10800 ");
            this.Formulas.Add("sum #1 0 10800 ");
            this.Formulas.Add("prod @7 #2 @6 ");
            this.Formulas.Add("sum 21600 0 @8");


            this.AdjustmentValues = "6480,8640,4320";
            this.ConnectorLocations = "Rectangle";

            this.TextboxRectangle = "@8,@1,@9,@4;@1,@8,@4,@9";
           
            this.Handles = new List<Handle>();

            var HandleOne = new Handle
            {
                position = "#0,topLeft",
                xrange = "@2,@1"
            };
            this.Handles.Add(HandleOne);
            var HandleTwo = new Handle
            {
                position = "#1,#2",
                xrange = "@0,10800",
                yrange = "0,@0"
            };
            this.Handles.Add(HandleTwo);

        }
    }
}
