using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(170)]
    public class TextFadeUp : ShapeType
    {
        public TextFadeUp()
        {
            this.TextPath = true;
            this.Joins = JoinStyle.none;
            this.AdjustmentValues = "7200";
            this.Path = "m@0,l@1,m,21600r21600,e";
            this.ConnectorLocations = "10800,0;@2,10800;10800,21600;@3,10800";
            this.ConnectorAngles = "270,180,90,0";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum 21600 0 @0");
            this.Formulas.Add("prod #0 1 2");
            this.Formulas.Add("sum 21600 0 @2");
            this.Formulas.Add("sum @1 21600 @0");

            this.Handles = new List<Handle>();
            var h1 = new Handle
            {
                position = "#0,topLeft",
                xrange = "0,10792"
            };
            this.Handles.Add(h1);

        }
    }
}
