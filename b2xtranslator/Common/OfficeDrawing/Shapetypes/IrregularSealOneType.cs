namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(71)]
    public class IrregularSealOneType : ShapeType
    {
        public IrregularSealOneType()
        {
            this.ShapeConcentricFill = false;
            
            this.Joins = JoinStyle.miter;

            this.Path = "m10800,5800l8352,2295,7312,6320,370,2295,4627,7617,,8615r3722,3160l135,14587r5532,-650l4762,17617,7715,15627r770,5973l10532,14935r2715,4802l14020,14457r4125,3638l16837,12942r4763,348l17607,10475,21097,8137,16702,7315,18380,4457r-4225,868l14522,xe";

            this.ConnectorLocations = "14522,0;0,8615;8485,21600;21600,13290";

            this.ConnectorAngles = "270,180,90,0";

            this.TextboxRectangle = "4627,6320,16702,13937";
        }
    }
}
