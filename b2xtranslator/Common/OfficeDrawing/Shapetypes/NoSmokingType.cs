using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(57)]
    class NoSmokingType : ShapeType
    {
        public NoSmokingType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m,10800qy10800,,21600,10800,10800,21600,,10800xar@0@0@16@16@12@14@15@13xar@0@0@16@16@13@15@14@12xe";
            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("prod @0 2 1");
            this.Formulas.Add("sum 21600 0 @1");
            this.Formulas.Add("prod @2 @2 1 ");
            this.Formulas.Add("prod @0 @0 1");
            this.Formulas.Add("sum @3 0 @4");
            this.Formulas.Add("prod @5 1 8 ");
            this.Formulas.Add("sqrt @6 ");
            this.Formulas.Add("prod @4 1 8 ");
            this.Formulas.Add("sqrt @8 ");
            this.Formulas.Add("sum @7 @9 0");
            this.Formulas.Add("sum @7 0 @9");
            this.Formulas.Add("sum @10 10800 0");
            this.Formulas.Add("sum 10800 0 @10");
            this.Formulas.Add("sum @11 10800 0 ");
            this.Formulas.Add("sum 10800 0 @11 ");
            this.Formulas.Add("sum 21600 0 @0");



            this.AdjustmentValues = "2700";
            this.ConnectorLocations = "10800,0;3163,3163;0,10800;3163,18437;10800,21600;18437,18437;21600,10800;18437,3163";

            this.TextboxRectangle = "3163,3163,18437,18437";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,center",
                xrange = "0,7200"
            };
            this.Handles.Add(HandleOne);

        }

    }
}
