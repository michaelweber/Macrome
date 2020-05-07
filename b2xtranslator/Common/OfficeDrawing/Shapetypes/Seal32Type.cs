using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(60)]
    class Seal32Type : ShapeType
    {
        public Seal32Type()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m21600,10800l@9@18,21392,8693@11@20,20777,6667@13@22,19780,4800@15@24,18436,3163@16@23,16800,1820@14@21,14932,822@12@19,12907,208@10@17,10800,0@18@17,8693,208@20@19,6667,822@22@21,4800,1820@24@23,3163,3163@23@24,1820,4800@21@22,822,6667@19@20,208,8693@17@18,,10800@17@10,208,12907@19@12,822,14932@21@14,1820,16800@23@16,3163,18436@24@15,4800,19780@22@13,6667,20777@20@11,8693,21392@18@9,10800,21600@10@9,12907,21392@12@11,14932,20777@14@13,16800,19780@16@15,18436,18436@15@16,19780,16800@13@14,20777,14932@11@12,21392,12907@9@10xe";

            this.Formulas = new List<string>();
            

            this.Formulas.Add("sum 10800 0 #0"); 
            this.Formulas.Add("prod @0 32610 32768"); 
            this.Formulas.Add("prod @0 3212 32768"); 
            this.Formulas.Add("prod @0 31357 32768"); 
            this.Formulas.Add("prod @0 9512 32768"); 
            this.Formulas.Add("prod @0 28899 32768"); 
            this.Formulas.Add("prod @0 15447 32768"); 
            this.Formulas.Add("prod @0 25330 32768"); 
            this.Formulas.Add("prod @0 20788 32768"); 
            this.Formulas.Add("sum @1 10800 0"); 
            this.Formulas.Add("sum @2 10800 0"); 
            this.Formulas.Add("sum @3 10800 0"); 
            this.Formulas.Add("sum @4 10800 0"); 
            this.Formulas.Add("sum @5 10800 0"); 
            this.Formulas.Add("sum @6 10800 0"); 
            this.Formulas.Add("sum @7 10800 0"); 
            this.Formulas.Add("sum @8 10800 0"); 
            this.Formulas.Add("sum 10800 0 @1"); 
            this.Formulas.Add("sum 10800 0 @2"); 
            this.Formulas.Add("sum 10800 0 @3"); 
            this.Formulas.Add("sum 10800 0 @4"); 
            this.Formulas.Add("sum 10800 0 @5"); 
            this.Formulas.Add("sum 10800 0 @6"); 
            this.Formulas.Add("sum 10800 0 @7"); 
            this.Formulas.Add("sum 10800 0 @8"); 
            this.Formulas.Add("prod @0 23170 32768"); 
            this.Formulas.Add("sum @25 10800 0");
            this.Formulas.Add("sum 10800 0 @25");


            this.AdjustmentValues = "2700";
            this.ConnectorLocations = "Rectangle";
            this.TextboxRectangle = "@27,@27,@26,@26";


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


