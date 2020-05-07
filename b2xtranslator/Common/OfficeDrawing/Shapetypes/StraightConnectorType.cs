namespace b2xtranslator.OfficeDrawing.Shapetypes
{
    [OfficeShapeType(32)]
    public class StraightConnectorType : ShapeType
    {
        public StraightConnectorType()
        {
            this.Path = "m,l21600,21600e";
            this.ConnectorLocations = "0,0;21600,21600";
        }
    }
}
