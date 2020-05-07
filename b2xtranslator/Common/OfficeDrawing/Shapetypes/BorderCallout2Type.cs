using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(48)]
    public class BorderCallout2Type : ShapeType
    {
        public BorderCallout2Type()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;


            this.Path = "m@0@1l@2@3@4@5nfem,l21600,r,21600l,21600xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("val #1");
            this.Formulas.Add("val #2");
            this.Formulas.Add("val #3");
            this.Formulas.Add("val #4");
            this.Formulas.Add("val #5");
            this.AdjustmentValues = "-10080,24300,-3600,4050,-1800,4050";
            this.ConnectorLocations = "@0,@1;10800,0;10800,21600;0,10800;21600,10800";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "#0,#1"
            };
            this.Handles.Add(HandleOne);

            var HandleTwo = new Handle
            {
                position = "#2,#3"
            };
            this.Handles.Add(HandleTwo);

            var HandleThree = new Handle
            {
                position = "#4,#5"
            };
            this.Handles.Add(HandleThree);
        }
    }
}
