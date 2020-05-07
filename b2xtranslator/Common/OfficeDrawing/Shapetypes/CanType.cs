using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(22)]
    public class CanType :ShapeType
    {
        public CanType()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.round;

            this.Path = "m10800,qx0@1l0@2qy10800,21600,21600@2l21600@1qy10800,xem0@1qy10800@0,21600@1nfe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("prod #0 1 2");
            this.Formulas.Add("sum height 0 @1");

            this.AdjustmentValues = "5400";

            this.ConnectorLocations = "10800,@0;10800,0;0,10800;10800,21600;21600,10800";

            this.ConnectorAngles = "270,270,180,90,0";

            this.TextboxRectangle = "0,@0,21600,@2";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "center,#0",
                yrange = "0,10800"
            };
            this.Handles.Add(HandleOne);
        }
    }
}
