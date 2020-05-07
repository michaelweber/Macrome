using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(91)]
    class BentArrowType : ShapeType
    {
        public BentArrowType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            this.Path = "m21600,6079l@0,0@0@1,12427@1qx,12158l,21600@4,21600@4,12158qy12427@2l@0@2@0,12158xe";
            this.Formulas = new List<string>();
            this.Formulas.Add("val #0 ");
            this.Formulas.Add("val #1 ");
            this.Formulas.Add("sum 12158 0 #1 ");
            this.Formulas.Add("sum @2 0 #1 ");
            this.Formulas.Add("prod @3 32768 32059 ");
            this.Formulas.Add("prod @4 1 2 ");
            this.Formulas.Add("sum 21600 0 #0 "); 
            this.Formulas.Add("prod @6 #1 6079 ");
            this.Formulas.Add("sum @7 #0 0");

            this.AdjustmentValues = "Connector Angles";
            this.ConnectorLocations = "@0,0;@0,12158;@5,21600;21600,6079";
            this.ConnectorAngles = "270,90,90,0";

            this.TextboxRectangle = "12427,@1,@8,@2;0,12158,@4,21600";
           
            this.Handles = new List<Handle>();

            var HandleOne = new Handle
            {
                position = "#0,#1",
                xrange = "12427,21600",
                yrange = "0,6079"
            };
            this.Handles.Add(HandleOne);

        }
    }
}
