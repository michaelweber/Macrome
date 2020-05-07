using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(152)]
    public class TextCurveUp : ShapeType
    {
        public TextCurveUp()
        {
            this.TextPath = true;
            this.Joins = JoinStyle.none;
            this.AdjustmentValues = "9931";
            this.Path = "m0@0c7200@2,14400@1,21600,m0@5c7200@6,14400@6,21600@5e";
            this.ConnectorLocations = "10800,@10;0,@9;10800,21600;21600,@8";
            this.ConnectorAngles = "270,180,90,0";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("prod #0 3 4");
            this.Formulas.Add("prod #0 5 4");
            this.Formulas.Add("prod #0 3 8");
            this.Formulas.Add("prod #0 1 8");
            this.Formulas.Add("sum 21600 0 @3");
            this.Formulas.Add("sum @4 21600 0");
            this.Formulas.Add("prod #0 1 2");
            this.Formulas.Add("prod @5 1 2");
            this.Formulas.Add("sum @7 @8 0");
            this.Formulas.Add("prod #0 7 8");
            this.Formulas.Add("prod @5 1 3");
            this.Formulas.Add("sum @1 @2 0");
            this.Formulas.Add("sum @12 @0 0");
            this.Formulas.Add("prod @13 1 4");
            this.Formulas.Add("sum @11 14400 @14");

            this.Handles = new List<Handle>();
            var h1 = new Handle
            {
                position = "topLeft,#0",
                yrange = "0,12169"
            };
            this.Handles.Add(h1);
        }
    }
}
