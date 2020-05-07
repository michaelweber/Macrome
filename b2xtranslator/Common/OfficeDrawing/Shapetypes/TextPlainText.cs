using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(136)]
    public class TextPlainText : ShapeType
    {
        public TextPlainText()
        {
            this.TextPath = true;

            this.Joins = JoinStyle.none;

            this.AdjustmentValues = "10800";

            this.Path = "m@7,l@8,m@5,21600l@6,21600e";

            this.Formulas = new List<string>();
            this.Formulas.Add("sum #0 0 10800");
            this.Formulas.Add("prod #0 2 1");
            this.Formulas.Add("sum 21600 0 @1");
            this.Formulas.Add("sum 0 0 @2");
            this.Formulas.Add("sum 21600 0 @3");
            this.Formulas.Add("if @0 @3 0");
            this.Formulas.Add("if @0 21600 @1");
            this.Formulas.Add("if @0 0 @2");
            this.Formulas.Add("if @0 @4 21600");
            this.Formulas.Add("mid @5 @6");
            this.Formulas.Add("mid @8 @5");
            this.Formulas.Add("mid @7 @8");
            this.Formulas.Add("mid @6 @7");
            this.Formulas.Add("sum @6 0 @5");

            this.ConnectorLocations = "@9,0;@10,10800;@11,21600;@12,10800";
            this.ConnectorAngles = "270,180,90,0";

            this.Handles = new List<Handle>();
            var h1 = new Handle
            {
                position = "#0,bottomRight",
                xrange = "6629,14971"
            };
            this.Handles.Add(h1);
        }
    }
}
