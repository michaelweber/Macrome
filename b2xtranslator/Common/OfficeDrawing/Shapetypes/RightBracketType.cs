using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(86)]
    public class RightBracketType :ShapeType
    {

        public RightBracketType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.round;
            //Endcaps: Flat

            this.Path = "m,qx21600@0l21600@1qy,21600e";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum 21600 0 #0");
            this.Formulas.Add("prod #0 9598 32768");
            this.Formulas.Add("sum 21600 0 @2");

            this.AdjustmentValues = "1800";
            this.ConnectorLocations = "0,0;0,21600;21600,10800";
            this.TextboxRectangle = "0,@2,15274,@3";


            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "bottomRight,#0",
                yrange = "0,10800"
            };
            this.Handles.Add(HandleOne);
        }
    }
}
