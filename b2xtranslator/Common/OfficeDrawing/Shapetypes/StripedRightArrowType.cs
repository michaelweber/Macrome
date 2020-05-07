using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(93)]
    class StripedRightArrowType : ShapeType
    {
        public StripedRightArrowType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m@0,l@0@1,3375@1,3375@2@0@2@0,21600,21600,10800xem1350@1l1350@2,2700@2,2700@1xem0@1l0@2,675@2,675@1xe";
            this.Formulas = new List<string>();
            
            this.Formulas.Add("val #0"); 
            this.Formulas.Add("val #1"); 
            this.Formulas.Add("sum height 0 #1"); 
            this.Formulas.Add("sum 10800 0 #1"); 
            this.Formulas.Add("sum width 0 #0"); 
            this.Formulas.Add("prod @4 @3 10800");
            this.Formulas.Add("sum width 0 @5");


            this.AdjustmentValues = "16200,5400";
            this.ConnectorLocations = "@0,0;0,10800;@0,21600;21600,10800";
            this.ConnectorAngles = "270,180,90,0"; 

            this.TextboxRectangle = "3375,@1,@6,@2";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,#1",
                xrange = "3375,21600",
                yrange = "0,10800"
            };

            this.Handles.Add(HandleOne);


        }
    }
}
