using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(172)]
    public class TextSlantUp : ShapeType
    {
        public TextSlantUp()
        {
            this.TextPath = true;

            this.Joins = JoinStyle.none;

            this.AdjustmentValues = "12000";

            this.Path = "m0@0l21600,m,21600l21600@1e";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum 21600 0 @0");
            this.Formulas.Add("prod #0 1 2");
            this.Formulas.Add("sum @2 10800 0");
            this.Formulas.Add("prod @1 1 2");
            this.Formulas.Add("sum @4 10800 0");

            this.ConnectorLocations = "10800,@2;0,@3;10800,@5;21600,@4";
            this.ConnectorAngles = "270,180,90,0";

            this.Handles = new List<Handle>();
            var h1 = new Handle
            {
                position = "topLeft,#0",
                yrange = "0,15429"
            };
            this.Handles.Add(h1);
        }
    }
}
