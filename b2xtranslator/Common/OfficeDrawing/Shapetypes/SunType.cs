using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{

    [OfficeShapeType(183)]
    class SunType : ShapeType
    {
        public SunType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m21600,10800l@15@14@15@18xem18436,3163l@17@12@16@13xem10800,l@14@10@18@10xem3163,3163l@12@13@13@12xem,10800l@10@18@10@14xem3163,18436l@13@16@12@17xem10800,21600l@18@15@14@15xem18436,18436l@16@17@17@16xem10800@19qx@19,10800,10800@20@20,10800,10800@19xe";
            this.Formulas = new List<string>();
            this.Formulas.Add("sum 10800 0 #0 ");
            this.Formulas.Add("prod @0 30274 32768 ");
            this.Formulas.Add("prod @0 12540 32768 ");
            this.Formulas.Add("sum @1 10800 0 ");
            this.Formulas.Add("sum @2 10800 0 ");
            this.Formulas.Add("sum 10800 0 @1 ");
            this.Formulas.Add("sum 10800 0 @2 ");
            this.Formulas.Add("prod @0 23170 32768 ");
            this.Formulas.Add("sum @7 10800 0 ");
            this.Formulas.Add("sum 10800 0 @7 ");
            this.Formulas.Add("prod @5 3 4 ");
            this.Formulas.Add("prod @6 3 4 ");
            this.Formulas.Add("sum @10 791 0 ");
            this.Formulas.Add("sum @11 791 0 ");
            this.Formulas.Add("sum @11 2700 0");
            this.Formulas.Add("sum 21600 0 @10 ");
            this.Formulas.Add("sum 21600 0 @12 ");
            this.Formulas.Add("sum 21600 0 @13 ");
            this.Formulas.Add("sum 21600 0 @14 ");
            this.Formulas.Add("val #0 ");
            this.Formulas.Add("sum 21600 0 #0");
            
            this.AdjustmentValues = "5400";
            this.TextboxRectangle = "@9,@9,@8,@8";
            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,center",
                xrange = "2700,10125"
            };
            this.Handles.Add(HandleOne);

        }
    }
}
