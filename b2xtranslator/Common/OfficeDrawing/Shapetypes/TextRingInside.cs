using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(142)]
    public class TextRingInside : ShapeType
    {
        public TextRingInside()
        {
            this.TextPath = true;
            this.Joins = JoinStyle.none;
            this.AdjustmentValues = "13500";
            this.Path = "m0@1qy10800,,21600@1,10800@0,0@1m0@2qy10800@3,21600@2,10800,21600,0@2e";
            this.ConnectorType = "custom";
            this.ConnectorLocations = "10800,0;10800,@0;0,10800;10800,21600;10800,@3;21600,10800";
            this.ConnectorAngles = "270,270,180,90,90,0";
            this.ExtrusionOk = true;
            this.Lock = new ProtectionBooleans
            {
                fUsefLockText = true,
                fLockText = true
            };
            this.LockShapeType = true;

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("prod #0 1 2");
            this.Formulas.Add("sum height 0 @1");
            this.Formulas.Add("sum height 0 #0");
            this.Formulas.Add("sum @2 0 @1");

            this.Handles = new List<Handle>();
            var h1 = new Handle
            {
                position = "center,#0",
                yrange = "10800,21600"
            };
            this.Handles.Add(h1);
        }
    }
}
