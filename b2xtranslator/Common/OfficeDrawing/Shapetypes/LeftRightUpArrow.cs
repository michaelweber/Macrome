using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(182)]
    class LeftRightUpArrow : ShapeType
    {
        public LeftRightUpArrow()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m10800,l@0@2@1@2@1@6@7@6@7@5,0@8@7,21600@7@9@10@9@10,21600,21600@8@10@5@10@6@4@6@4@2@3@2xe";
            this.Formulas = new List<string>();
            this.Formulas.Add("val #0 ");
            this.Formulas.Add("val #1 ");
            this.Formulas.Add("val #2 ");
            this.Formulas.Add("sum 21600 0 #0");
            this.Formulas.Add("sum 21600 0 #1");
            this.Formulas.Add("prod @0 21600 @3 ");
            this.Formulas.Add("prod @1 21600 @3 ");
            this.Formulas.Add("prod @2 @3 21600 ");
            this.Formulas.Add("prod 10800 21600 @3 ");
            this.Formulas.Add("prod @4 21600 @3 ");
            this.Formulas.Add("sum 21600 0 @7 ");
            this.Formulas.Add("sum @5 0 @8 ");
            this.Formulas.Add("sum @6 0 @8 ");
            this.Formulas.Add("prod @12 @7 @11 ");
            this.Formulas.Add("sum 21600 0 @13 ");
            this.Formulas.Add("sum @0 0 10800 ");
            this.Formulas.Add("sum @1 0 10800 ");
            this.Formulas.Add("prod @2 @16 @15");



            this.AdjustmentValues = "6480,8640,6171";
            this.ConnectorLocations = "10800,0;0,@8;10800,@9;21600,@8";
            this.ConnectorAngles = "270,180,90,0";

            this.TextboxRectangle = "@13,@6,@14,@9;@1,@17,@4,@9";
           
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
                yrange = "0,@5"
            };
            this.Handles.Add(HandleTwo);

        }
    }
}
