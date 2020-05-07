using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(87)]
    public class LeftBraceType : ShapeType
    {
        public LeftBraceType()
        {
            this.ShapeConcentricFill = false;
            this.Joins = JoinStyle.miter;
            //Endcaps: Flat

            this.Path = "m21600,qx10800@0l10800@2qy0@11,10800@3l10800@1qy21600,21600e";
            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum 21600 0 #0");
            this.Formulas.Add("sum #1 0 #0");
            this.Formulas.Add("sum #1 #0 0");
            this.Formulas.Add("prod #0 9598 32768");
            this.Formulas.Add("sum 21600 0 @4");
            this.Formulas.Add("sum 21600 0 #1");
            this.Formulas.Add("min #1 @6");
            this.Formulas.Add("prod @7 1 2");
            this.Formulas.Add("prod #0 2 1");
            this.Formulas.Add("sum 21600 0 @9");
            this.Formulas.Add("val #1");
            
            this.AdjustmentValues = "1800,10800";
            this.ConnectorLocations = "21600,0;0,10800;21600,21600";
            this.TextboxRectangle = "13963,@4,21600,@5";

            this.Handles = new List<Handle>();
            var HandleOne = new Handle();
            var HandleTwo= new Handle();
            HandleOne.position="center,#0";
            HandleOne.yrange = "0,@8";
            HandleTwo.position="topLeft,#1";
            HandleTwo.yrange="@9,@10";
            this.Handles.Add(HandleOne);
            this.Handles.Add(HandleTwo);


        }
    }
}
