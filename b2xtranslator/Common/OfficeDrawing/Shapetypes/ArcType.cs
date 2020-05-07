using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(19)]
    class ArcType : ShapeType
    {
        public ArcType()
        {
            this.ShapeConcentricFill = true;
            this.Joins = JoinStyle.round;
            this.Path = "wr-21600,,21600,43200,,,21600,21600nfewr-21600,,21600,43200,,,21600,21600l,21600nsxe";
            this.Formulas = new List<string>();
              
            this.Formulas.Add("val #2");
            this.Formulas.Add("val #3");
            this.Formulas.Add("val #4");

            this.AdjustmentValues = "-5898240,,,21600,21600";
            this.ConnectorLocations = "0,0;21600,21600;0,21600";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle
            {
                position = "@2,#0",
                polar = "@0,@1"
            };
            this.Handles.Add(HandleOne);

            var HandleTwo = new Handle();
            HandleOne.position = "@2,#1";
            HandleOne.polar = "@0,@1"; 
            this.Handles.Add(HandleTwo);
        }
    }
}
