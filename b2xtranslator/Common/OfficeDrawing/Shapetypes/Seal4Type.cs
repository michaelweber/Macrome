using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(187)]
    class Seal4Type : ShapeType
    {
        public Seal4Type()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m21600,10800l@2@3,10800,0@3@3,,10800@3@2,10800,21600@2@2xe";
            
            this.Formulas = new List<string>();
            this.Formulas.Add("sum 10800 0 #0");     
            this.Formulas.Add("prod @0 23170 32768");     
            this.Formulas.Add("sum @1 10800 0");     
            this.Formulas.Add("sum 10800 0 @1");

            this.AdjustmentValues = "8100"; 
            this.ConnectorLocations = "Rectangle";
            this.TextboxRectangle = "@3,@3,@2,@2";

            
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


