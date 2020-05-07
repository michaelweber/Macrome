using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(150)]
    public class TextCirclePour : ShapeType
    {
        public TextCirclePour()
        {
            this.TextPath = true;
            this.Joins = JoinStyle.none;
            this.AdjustmentValues = "-11730944,5400";
            this.Path = "al10800,10800,10800,10800@2@5al10800,10800@0@0@2@5e";
            this.ConnectorLocations = "@17,10800;@12,@13;@16,10800;@12,@14";
            this.ExtrusionOk = true;
            this.Lock = new ProtectionBooleans
            {
                fUsefLockText = true,
                fLockText = true
            };
            this.LockShapeType = true;

            this.Formulas = new List<string>();
            this.Formulas.Add("val #1");
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum 0 0 #0");
            this.Formulas.Add("prod #0 2 1");
            this.Formulas.Add("sumangle @3 0 360");
            this.Formulas.Add("if @3 @4 @3");
            this.Formulas.Add("val 10800");
            this.Formulas.Add("sum 10800 0 #1");
            this.Formulas.Add("prod #1 1 2");
            this.Formulas.Add("sum @8 5400 0");
            this.Formulas.Add("cos @9 #0");
            this.Formulas.Add("sin @9 #0");
            this.Formulas.Add("sum @10 10800 0");
            this.Formulas.Add("sum @11 10800 0");
            this.Formulas.Add("sum 10800 0 @11");
            this.Formulas.Add("sum #1 10800 0");
            this.Formulas.Add("if #0 @7 @15");
            this.Formulas.Add("if #0 0 21600");

            this.Handles = new List<Handle>();
            var h1 = new Handle
            {
                position = "#1,#0",
                polar = "10800,10800",
                radiusrange = "0,10800"
            };
            this.Handles.Add(h1);

        }
    }
}
