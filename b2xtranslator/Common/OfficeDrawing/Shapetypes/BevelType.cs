using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(84)]
    public class BevelType : ShapeType
    {
        public BevelType()
        {
            this.ShapeConcentricFill = true;
            
            this.Joins = JoinStyle.miter;

            this.Path = "m,l,21600r21600,l21600,xem@0@0nfl@0@2@1@2@1@0xem,nfl@0@0em,21600nfl@0@2em21600,21600nfl@1@2em21600,nfl@1@0e";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum width 0 #0");
            this.Formulas.Add("sum height 0 #0");
            this.Formulas.Add("prod width 1 2");
            this.Formulas.Add("prod height 1 2");
            this.Formulas.Add("prod #0 1 2");
            this.Formulas.Add("prod #0 3 2");
            this.Formulas.Add("sum @1 @5 0");
            this.Formulas.Add("sum @2 @5 0");

            this.AdjustmentValues = "2700";
            
            this.ConnectorLocations = "0,@4;@0,@4;@3,21600;@3,@2;21600,@4;@1,@4;@3,0;@3,@0";

            this.TextboxRectangle = "@0,@0,@1,@2";
            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,topLeft",
                switchHandle = "true",
                xrange = "0,10800"
            };
            this.Handles.Add(HandleOne);

            this.Limo = "10800,10800";
        }
    }
}
