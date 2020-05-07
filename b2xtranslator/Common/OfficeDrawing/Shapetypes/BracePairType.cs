using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(186)]
    public class BracePairType : ShapeType
    {
        public BracePairType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            //Encaps: Flat

            this.Path = "m@9,nfqx@0@0l@0@7qy0@4@0@8l@0@6qy@9,21600em@10,nfqx@5@0l@5@7qy21600@4@5@8l@5@6qy@10,21600em@9,nsqx@0@0l@0@7qy0@4@0@8l@0@6qy@9,21600l@10,21600qx@5@6l@5@8qy21600@4@5@7l@5@0qy@10,xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("val width");
            this.Formulas.Add("val height");
            this.Formulas.Add("prod width 1 2");
            this.Formulas.Add("prod height 1 2");
            this.Formulas.Add("sum width 0 #0");
            this.Formulas.Add("sum height 0 #0");
            this.Formulas.Add("sum @4 0 #0");
            this.Formulas.Add("sum @4 #0 0");
            this.Formulas.Add("prod #0 2 1");
            this.Formulas.Add("sum width 0 @9");
            this.Formulas.Add("prod #0 9598 32768");
            this.Formulas.Add("sum height 0 @11");
            this.Formulas.Add("sum @11 #0 0");
            this.Formulas.Add("sum width 0 @13");

            this.AdjustmentValues = "1800";
            this.ConnectorLocations = "@3,0;0,@4;@3,@2;@1,@4";
            this.TextboxRectangle = "@13,@11,@14,@12";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "topLeft,#0",
                switchHandle = "true",
                yrange = "0,5400"
            };
            this.Handles.Add(HandleOne);

            this.Limo = "10800,10800";


        }
    }
}
