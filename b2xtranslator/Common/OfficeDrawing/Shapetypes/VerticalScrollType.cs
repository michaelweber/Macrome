using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(97)]
    public class VerticalScrollType : ShapeType
    {
        public VerticalScrollType()
        {
            this.ShapeConcentricFill = false;
            
            this.Joins = JoinStyle.miter;

            this.Path = "m@5,qx@1@2l@1@0@2@0qx0@7@2,21600l@9,21600qx@10@7l@10@1@11@1qx21600@2@11,xem@5,nfqx@6@2@5@1@4@3@5@2l@6@2em@5@1nfl@10@1em@2,21600nfqx@1@7l@1@0em@2@0nfqx@3@8@2@7l@1@7e";

            this.AdjustmentValues="2700"; 
            this.ConnectorLocations = "@14,0;@1,@13;@14,@12;@10,@13";

            this.ConnectorAngles = "270,180,90,0";

            this.TextboxRectangle = "@1,@1,@10,@7";

            this.Formulas = new List<string>();
            this.Formulas.Add("sum height 0 #0 ");
            this.Formulas.Add("val #0 ");
            this.Formulas.Add("prod @1 1 2 ");
            this.Formulas.Add("prod @1 3 4 ");
            this.Formulas.Add("prod @1 5 4 ");
            this.Formulas.Add("prod @1 3 2 ");
            this.Formulas.Add("prod @1 2 1 ");
            this.Formulas.Add("sum height 0 @2 ");
            this.Formulas.Add("sum height 0 @3 ");
            this.Formulas.Add("sum width 0 @5 ");
            this.Formulas.Add("sum width 0 @1 ");
            this.Formulas.Add("sum width 0 @2"); 
            this.Formulas.Add("val height ");
            this.Formulas.Add("prod height 1 2"); 
            this.Formulas.Add("prod width 1 2");

            this.Handles = new List<Handle>();
            var handleOne = new Handle
            {
                position = "topLeft,#0",
                yrange = "0,5400"
            };

            this.Handles.Add(handleOne); 
            this.Limo="10800,10800"; 
        }
    }
}
