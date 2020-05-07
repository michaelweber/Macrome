using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(52)]
    public class AccentBorderCallout3Type : ShapeType
    {
        public AccentBorderCallout3Type()
        {
            this.ShapeConcentricFill = true;

            this.Joins = JoinStyle.miter;

            this.Path = "m@0@1l@2@3@4@5@6@7nfem@6,l@6,21600nfem,l21600,r,21600l,21600xe";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("val #1");
            this.Formulas.Add("val #2");
            this.Formulas.Add("val #3");
            this.Formulas.Add("val #4");
            this.Formulas.Add("val #5");
            this.Formulas.Add("val #6");
            this.Formulas.Add("val #7");
            this.AdjustmentValues = "23400,24400,25200,21600,25200,4050,23400,4050";
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

            var HandleFour = new Handle
            {
                position = "#6,#7"
            };
            this.Handles.Add(HandleFour);
        }
    }
}
