using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(98)]
    public class HorizontalScrollType : ShapeType
    {
        public HorizontalScrollType()
        {
            this.ShapeConcentricFill = false;

            this.Joins = JoinStyle.miter;

            this.Path = "m0@5qy@2@1l@0@1@0@2qy@7,,21600@2l21600@9qy@7@10l@1@10@1@11qy@2,21600,0@11xem0@5nfqy@2@6@1@5@3@4@2@5l@2@6em@1@5nfl@1@10em21600@2nfqy@7@1l@0@1em@0@2nfqy@8@3@7@2l@7@1e";

            this.AdjustmentValues = "2700";
            this.ConnectorLocations = "@13,@1;0,@14;@13,@10;@12,@14";

            this.ConnectorAngles = "270,180,90,0";

            this.TextboxRectangle = "@1,@1,@7,@10";

            this.Formulas = new List<string>();


            this.Formulas.Add("sum width 0 #0"); 
            this.Formulas.Add("val #0 ");
            this.Formulas.Add("prod @1 1 2"); 
            this.Formulas.Add("prod @1 3 4 ");
            this.Formulas.Add("prod @1 5 4 ");
            this.Formulas.Add("prod @1 3 2 ");
            this.Formulas.Add("prod @1 2 1 ");
            this.Formulas.Add("sum width 0 @2 ");
            this.Formulas.Add("sum width 0 @3 ");
            this.Formulas.Add("sum height 0 @5 ");
            this.Formulas.Add("sum height 0 @1 ");
            this.Formulas.Add("sum height 0 @2 ");
            this.Formulas.Add("val width ");
            this.Formulas.Add("prod width 1 2"); 
            this.Formulas.Add("prod height 1 2");

            this.Handles = new List<Handle>();
            var handleOne = new Handle
            {
                position = "#0,topLeft",
                xrange = "0,5400"
            };

            this.Handles.Add(handleOne);
            this.Limo = "10800,10800"; 
        }
    }
}
