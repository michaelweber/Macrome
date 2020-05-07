using System.Collections.Generic;

namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(161)]
    public class TextDeflate : ShapeType
    {
        public TextDeflate()
        {
            this.TextPath = true;

            this.Path = "m,c7200@0,14400@0,21600,m,21600c7200@1,14400@1,21600,21600e";

            this.Formulas = new List<string>();
            this.Formulas.Add("prod #0 4 3");
            this.Formulas.Add("sum 21600 0 @0");
            this.Formulas.Add("val #0");
            this.Formulas.Add("sum 21600 0 #0");

            this.ConnectorLocations = "10800,@2;0,10800;10800,@3;21600,10800";
            this.ConnectorAngles = "270,180,90,0";
                
            this.Handles = new List<Handle>();
            var h1 = new Handle
            {
                position = "center,#0",
                yrange = "0,8100"
            };
            this.Handles.Add(h1);
        }
    }
}
