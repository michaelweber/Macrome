using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(58)]
    class Seal8Type : ShapeType
    {
        public Seal8Type()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m21600,10800l@3@6,18436,3163@4@5,10800,0@6@5,3163,3163@5@6,,10800@5@4,3163,18436@6@3,10800,21600@4@3,18436,18436@3@4xe";

            this.Formulas = new List<string>();



            this.Formulas.Add("sum 10800 0 #0"); 
            this.Formulas.Add("prod @0 30274 32768"); 
            this.Formulas.Add("prod @0 12540 32768"); 
            this.Formulas.Add("sum @1 10800 0"); 
            this.Formulas.Add("sum @2 10800 0"); 
            this.Formulas.Add("sum 10800 0 @1"); 
            this.Formulas.Add("sum 10800 0 @2"); 
            this.Formulas.Add("prod @0 23170 32768"); 
            this.Formulas.Add("sum @7 10800 0");
            this.Formulas.Add("sum 10800 0 @7");


            this.AdjustmentValues = "2538";
            this.ConnectorLocations = "Rectangle";
            this.TextboxRectangle = "@9,@9,@8,@8";


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


