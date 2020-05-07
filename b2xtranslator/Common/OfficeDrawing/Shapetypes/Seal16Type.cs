using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(59)]
    class Seal16Type : ShapeType
    {
        public Seal16Type()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m21600,10800l@5@10,20777,6667@7@12,18436,3163@8@11,14932,822@6@9,10800,0@10@9,6667,822@12@11,3163,3163@11@12,822,6667@9@10,,10800@9@6,822,14932@11@8,3163,18436@12@7,6667,20777@10@5,10800,21600@6@5,14932,20777@8@7,18436,18436@7@8,20777,14932@5@6xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("sum 10800 0 #0"); 
            this.Formulas.Add("prod @0 32138 32768"); 
            this.Formulas.Add("prod @0 6393 32768"); 
            this.Formulas.Add("prod @0 27246 32768"); 
            this.Formulas.Add("prod @0 18205 32768"); 
            this.Formulas.Add("sum @1 10800 0"); 
            this.Formulas.Add("sum @2 10800 0"); 
            this.Formulas.Add("sum @3 10800 0"); 
            this.Formulas.Add("sum @4 10800 0"); 
            this.Formulas.Add("sum 10800 0 @1"); 
            this.Formulas.Add("sum 10800 0 @2"); 
            this.Formulas.Add("sum 10800 0 @3"); 
            this.Formulas.Add("sum 10800 0 @4"); 
            this.Formulas.Add("prod @0 23170 32768"); 
            this.Formulas.Add("sum @13 10800 0"); 
            this.Formulas.Add("sum 10800 0 @13");

            this.AdjustmentValues = "2700";
            this.ConnectorLocations = "Rectangle";
            this.TextboxRectangle = "@15,@15,@14,@14";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,center",
                xrange = "0,10800"
            };


            this.Handles.Add(HandleOne);
        }
    }
}


