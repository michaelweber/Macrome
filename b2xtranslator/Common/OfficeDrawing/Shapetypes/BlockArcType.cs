using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(95)]
    class BlockArcType : ShapeType
    {
        public BlockArcType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "al10800,10800@0@0@2@14,10800,10800,10800,10800@3@15xe";
            this.Formulas = new List<string>();
            this.Formulas.Add("val #1 ");
            this.Formulas.Add("val #0 ");
            this.Formulas.Add("sum 0 0 #0 ");
            this.Formulas.Add("sumangle #0 0 180 ");
            this.Formulas.Add("sumangle #0 0 90 ");
            this.Formulas.Add("prod @4 2 1 ");
            this.Formulas.Add("sumangle #0 90 0 ");
            this.Formulas.Add("prod @6 2 1 ");
            this.Formulas.Add("abs #0 ");
            this.Formulas.Add("sumangle @8 0 90 ");
            this.Formulas.Add(" if @9 @7 @5 ");
            this.Formulas.Add("sumangle @10 0 360 ");
            this.Formulas.Add("if @10 @11 @10 ");
            this.Formulas.Add("sumangle @12 0 360 ");
            this.Formulas.Add("if @12 @13 @12 ");
            this.Formulas.Add("sum 0 0 @14 ");
            this.Formulas.Add("val 10800 ");
            this.Formulas.Add("sum 10800 0 #1 ");
            this.Formulas.Add("prod #1 1 2 ");
            this.Formulas.Add("sum @18 5400 0 ");
            this.Formulas.Add("cos @19 #0 ");
            this.Formulas.Add("sin @19 #0 ");
            this.Formulas.Add("sum @20 10800 0 ");
            this.Formulas.Add("sum @21 10800 0 ");
            this.Formulas.Add("sum 10800 0 @20 ");
            this.Formulas.Add("sum #1 10800 0 ");
            this.Formulas.Add("if @9 @17 @25 ");
            this.Formulas.Add("if @9 0 21600 ");
            this.Formulas.Add("cos 10800 #0 ");
            this.Formulas.Add("sin 10800 #0 ");
            this.Formulas.Add("sin #1 #0 ");
            this.Formulas.Add("sum @28 10800 0 ");
            this.Formulas.Add("sum @29 10800 0 ");
            this.Formulas.Add("sum @30 10800 0 ");
            this.Formulas.Add("if @4 0 @31 ");
            this.Formulas.Add("if #0 @34 0 ");
            this.Formulas.Add("if @6 @35 @31 ");
            this.Formulas.Add("sum 21600 0 @36 ");
            this.Formulas.Add("if @4 0 @33 ");
            this.Formulas.Add("if #0 @38 @32 ");
            this.Formulas.Add("if @6 @39 0 ");
            this.Formulas.Add("if @4 @32 21600");
            this.Formulas.Add("if @6 @41 @33");

            this.AdjustmentValues = "11796480,5400";
            this.ConnectorLocations = "10800,@27;@22,@23;10800,@26;@24,@23";

            this.TextboxRectangle = "@36,@40,@37,@42";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#1,#0",
                polar = "10800,10800",
                radiusrange = "0,10800",
                switchHandle = "true",
                xrange = "0,10800"
            };
            this.Handles.Add(HandleOne);


        }
    }
}
