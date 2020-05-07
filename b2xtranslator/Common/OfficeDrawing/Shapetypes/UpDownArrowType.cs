using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(70)]
    public class UpDownArrowType :ShapeType
    {
        public UpDownArrowType()
        {
            this.ShapeConcentricFill = false;

            this.Joins = JoinStyle.miter;
            this.Path = "m10800,l21600@0@3@0@3@2,21600@2,10800,21600,0@2@1@2@1@0,0@0xe";

            this.Formulas= new List<string>();
            this.Formulas.Add("val #1");
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum 21600 0 #1");
            this.Formulas.Add("sum 21600 0 #0");
            this.Formulas.Add("prod #1 #0 10800 ");
            this.Formulas.Add("sum #1 0 @4");
            this.Formulas.Add("sum 21600 0 @5");

            this.AdjustmentValues = "5400,4320";

            this.ConnectorLocations="10800,0;0,@0;@1,10800;0,@2;10800,21600;21600,@2;@3,10800;21600,@0";

            this.ConnectorAngles="270,180,180,180,90,0,0,0";

            this.TextboxRectangle="@1,@5,@3,@6";
            
            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,#1",
                xrange = "0,10800",
                yrange = "0,10800"
            };
            this.Handles.Add(HandleOne);

        }
    }
}
