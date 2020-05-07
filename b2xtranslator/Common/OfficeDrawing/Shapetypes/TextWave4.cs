using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(159)]
    public class TextWave4 : ShapeType
    {
        public TextWave4()
        {
            this.TextPath = true;
            this.Joins = JoinStyle.none;
            this.ExtrusionOk = true;
            this.Lock = new ProtectionBooleans
            {
                fUsefLockText = true,
                fLockText = true
            };
            this.LockShapeType = true;

            this.AdjustmentValues = "1404,10800";
            this.Path = "m@37@0c@38@1@39@3@40@0@41@1@42@3@43@0m@30@4c@31@6@32@5@33@4@34@6@35@5@36@4e";
            this.ConnectorLocations = "@40,@0;@51,10800;@33,@4;@50,10800";
            this.ConnectorAngles = "270,180,90,0";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("prod @0 41 9");
            this.Formulas.Add("prod @0 23 9");
            this.Formulas.Add("sum 0 0 @2");
            this.Formulas.Add("sum 21600 0 #0");
            this.Formulas.Add("sum 21600 0 @1");
            this.Formulas.Add("sum 21600 0 @3");
            this.Formulas.Add("sum #1 0 10800");
            this.Formulas.Add("sum 21600 0 #1");
            this.Formulas.Add("prod @8 1 3");
            this.Formulas.Add("prod @8 2 3");
            this.Formulas.Add("prod @8 4 3");
            this.Formulas.Add("prod @8 5 3");
            this.Formulas.Add("prod @8 2 1");
            this.Formulas.Add("sum 21600 0 @9");
            this.Formulas.Add("sum 21600 0 @10");
            this.Formulas.Add("sum 21600 0 @8");
            this.Formulas.Add("sum 21600 0 @11");
            this.Formulas.Add("sum 21600 0 @12");
            this.Formulas.Add("sum 21600 0 @13");
            this.Formulas.Add("prod #1 1 3");
            this.Formulas.Add("prod #1 2 3");
            this.Formulas.Add("prod #1 4 3");
            this.Formulas.Add("prod #1 5 3");
            this.Formulas.Add("prod #1 2 1");
            this.Formulas.Add("sum 21600 0 @20");
            this.Formulas.Add("sum 21600 0 @21");
            this.Formulas.Add("sum 21600 0 @22");
            this.Formulas.Add("sum 21600 0 @23");
            this.Formulas.Add("sum 21600 0 @24");
            this.Formulas.Add("if @7 @19 0");
            this.Formulas.Add("if @7 @18 @20");
            this.Formulas.Add("if @7 @17 @21");
            this.Formulas.Add("if @7 @16 #1");
            this.Formulas.Add("if @7 @15 @22");
            this.Formulas.Add("if @7 @14 @23");
            this.Formulas.Add("if @7 21600 @24");
            this.Formulas.Add("if @7 0 @29");
            this.Formulas.Add("if @7 @9 @28");
            this.Formulas.Add("if @7 @10 @27");
            this.Formulas.Add("if @7 @8 @8");
            this.Formulas.Add("if @7 @11 @26");
            this.Formulas.Add("if @7 @12 @25");
            this.Formulas.Add("if @7 @13 21600");
            this.Formulas.Add("sum @36 0 @30");
            this.Formulas.Add("sum @4 0 @0");
            this.Formulas.Add("max @30 @37");
            this.Formulas.Add("min @36 @43");
            this.Formulas.Add("prod @0 2 1");
            this.Formulas.Add("sum 21600 0 @48");
            this.Formulas.Add("mid @36 @43");
            this.Formulas.Add("mid @30 @37");


            this.Handles = new List<Handle>();
            var h1 = new Handle
            {
                position = "topLeft,#0",
                yrange = "0,2229"
            };
            this.Handles.Add(h1);
            var h2 = new Handle
            {
                position = "#1,bottomRight",
                xrange = "8640,12960"
            };
            this.Handles.Add(h2);
        }
    }
}
