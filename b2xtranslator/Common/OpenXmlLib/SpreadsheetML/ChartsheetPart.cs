using b2xtranslator.OpenXmlLib.DrawingML;

namespace b2xtranslator.OpenXmlLib.SpreadsheetML
{

    public class ChartsheetPart : OpenXmlPart
    {
        private DrawingsPart _drawingsPart = null;
        
        public ChartsheetPart(WorkbookPart parent, int partIndex)
            : base(parent, partIndex)
        {
        }


        public override string ContentType
        {
            get { return SpreadsheetMLContentTypes.Chartsheet; }
        }

        public override string RelationshipType
        {
            get { return OpenXmlRelationshipTypes.Chartsheet; }
        }

        public override string TargetName { get { return "sheet" + this.PartIndex.ToString(); } }
        public override string TargetDirectory { get { return "chartsheets"; } }

        public DrawingsPart DrawingsPart
        {
            get
            {
                if (this._drawingsPart == null)
                {
                    this._drawingsPart = this.AddPart(new DrawingsPart(this, ++((WorkbookPart)this.Parent).DrawingsNumber));
                    //this._drawingsPart = ((WorkbookPart)this.Parent).AddDrawingsPart();
                }
                return this._drawingsPart;
            }
        }
    }
}
