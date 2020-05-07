using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(144)]
    public class TextArchUpCurve : ShapeType
    {
        public TextArchUpCurve()
        {
            this.TextPath = true;
            this.Joins = JoinStyle.none;
            this.AdjustmentValues = "11796480";
            this.Path = "al10800,10800,10800,10800@2@14e";
            this.ConnectorLocations = "10800,@22;@19,@20;@21,@20";
            this.PreferRelative = false;
            this.TextKerning = true;
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
            this.Formulas.Add("sumangle #0 0 180");
            this.Formulas.Add("sumangle #0 0 90");
            this.Formulas.Add("prod @4 2 1");
            this.Formulas.Add("sumangle #0 90 0");
            this.Formulas.Add("prod @6 2 1");
            this.Formulas.Add("abs #0");
            this.Formulas.Add("sumangle @8 0 90");
            this.Formulas.Add("if @9 @7 @5");
            this.Formulas.Add("sumangle @10 0 360");
            this.Formulas.Add("if @10 @11 @10");
            this.Formulas.Add("sumangle @12 0 360");
            this.Formulas.Add("if @12 @13 @12");
            this.Formulas.Add("sum 0 0 @14");
            this.Formulas.Add("val 10800");
            this.Formulas.Add("cos 10800 #0");
            this.Formulas.Add("sin 10800 #0");
            this.Formulas.Add("sum @17 10800 0");
            this.Formulas.Add("sum @18 10800 0");
            this.Formulas.Add("sum 10800 0 @17");
            this.Formulas.Add("if @9 0 21600");
            this.Formulas.Add("sum 10800 0 @18");

            this.Handles = new List<Handle>();
            var h1 = new Handle
            {
                polar = "10800,10800",
                position = "@16,#0"
            };
            this.Handles.Add(h1);
        }
    }
}
