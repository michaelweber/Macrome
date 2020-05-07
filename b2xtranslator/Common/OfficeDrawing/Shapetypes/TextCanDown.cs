using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(175)]
    public class TextCanDown : ShapeType
    {
        public TextCanDown()
        {
            this.TextPath = true;

            this.Path = "m,qy10800@0,21600,m0@1qy10800,21600,21600@1e";

            this.Formulas = new List<string>();
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum 21600 0 #0");
            this.Formulas.Add("prod @1 1 2");
            this.Formulas.Add("sum @2 10800 0");

            this.ConnectorLocations = "10800,@0;0,@2;10800,21600;21600,@2";
            this.ConnectorAngles = "270,180,90,0";

            this.Handles = new List<Handle>();
            var h1 = new Handle
            {
                position = "center,#0",
                yrange = "0,7200"
            };
            this.Handles.Add(h1);

            
        }
    }
}
