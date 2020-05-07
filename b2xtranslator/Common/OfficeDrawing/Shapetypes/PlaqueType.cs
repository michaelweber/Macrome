using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(21)]
    class PlaqueType : ShapeType
    {
        public PlaqueType()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.miter;
            this.Path = "m@0,qy0@0l0@2qx@0,21600l@1,21600qy21600@2l21600@0qx@1,xe";
            this.Formulas = new List<string>();
            this.Formulas.Add("val #0 ");
            this.Formulas.Add("sum width 0 #0 ");
            this.Formulas.Add("sum height 0 #0 ");
            this.Formulas.Add("prod @0 7071 10000 ");
            this.Formulas.Add("sum width 0 @3 ");
            this.Formulas.Add("sum height 0 @3 ");
            this.Formulas.Add("val width ");
            this.Formulas.Add("val height ");
            this.Formulas.Add("prod width 1 2 ");
            this.Formulas.Add("prod height 1 2");


            
            this.AdjustmentValues = "3600";
            this.ConnectorLocations = "@8,0;0,@9;@8,@7;@6,@9";

            this.TextboxRectangle = "@3,@3,@4,@5";
            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,topLeft",
                switchHandle = "true",
                xrange = "0,10800"
            };
            this.Handles.Add(HandleOne);

            this.Limo = "10800,10800";
        }
    }
}
