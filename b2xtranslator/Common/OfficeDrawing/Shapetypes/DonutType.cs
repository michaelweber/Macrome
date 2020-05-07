using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(23)]
    class DonutType: ShapeType
    {
        public DonutType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.round;
            this.Path = "m,10800qy10800,,21600,10800,10800,21600,,10800xm@0,10800qy10800@2@1,10800,10800@0@0,10800xe";
                       
            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum width 0 #0");
            this.Formulas.Add("sum height 0 #0");
            this.Formulas.Add("prod @0 2929 10000");
            this.Formulas.Add("sum width 0 @3");
            this.Formulas.Add("sum height 0 @3");
            this.AdjustmentValues = "5400";
            this.ConnectorLocations = "10800,0;3163,3163;0,10800;3163,18437;10800,21600;18437,18437;21600,10800;18437,3163";
            this.TextboxRectangle = "3163,3163,18437,18437";
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
