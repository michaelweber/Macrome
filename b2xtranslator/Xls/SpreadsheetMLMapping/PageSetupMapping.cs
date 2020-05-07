

using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.SpreadsheetML;
using System.Globalization;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class PageSetupMapping : AbstractOpenXmlMapping,
          IMapping<PageSetupSequence>
    {
        ExcelContext _xlsContext;
        ChartsheetPart _chartsheetPart;

        public PageSetupMapping(ExcelContext xlsContext, ChartsheetPart targetPart)
            : base(targetPart.XmlWriter)
        {
            this._xlsContext = xlsContext;
            this._chartsheetPart = targetPart; 
        }

        #region IMapping<PageSetupSequence> Members

        public void Apply(PageSetupSequence pageSetupSequence)
        {
            // page margins
            this._writer.WriteStartElement(Sml.Sheet.ElPageMargins, Sml.Ns);
            {
                double leftMargin = pageSetupSequence.LeftMargin != null ? pageSetupSequence.LeftMargin.value : 0.75;
                double rightMargin = pageSetupSequence.RightMargin != null ? pageSetupSequence.RightMargin.value : 0.75;
                double topMargin = pageSetupSequence.TopMargin != null ? pageSetupSequence.TopMargin.value : 1.0;
                double bottomMargin = pageSetupSequence.BottomMargin != null ? pageSetupSequence.BottomMargin.value : 1.0;

                this._writer.WriteAttributeString(Sml.Sheet.AttrLeft, leftMargin.ToString(CultureInfo.InvariantCulture));
                this._writer.WriteAttributeString(Sml.Sheet.AttrRight, rightMargin.ToString(CultureInfo.InvariantCulture));
                this._writer.WriteAttributeString(Sml.Sheet.AttrTop, topMargin.ToString(CultureInfo.InvariantCulture));
                this._writer.WriteAttributeString(Sml.Sheet.AttrBottom, bottomMargin.ToString(CultureInfo.InvariantCulture));

                this._writer.WriteAttributeString(Sml.Sheet.AttrHeader, pageSetupSequence.Setup.numHdr.ToString(CultureInfo.InvariantCulture));
                this._writer.WriteAttributeString(Sml.Sheet.AttrFooter, pageSetupSequence.Setup.numFtr.ToString(CultureInfo.InvariantCulture));
            }
            this._writer.WriteEndElement();


            // page setup
            if (pageSetupSequence.Setup != null)
            {
                this._writer.WriteStartElement(Sml.Sheet.ElPageSetup, Sml.Ns);

                this._writer.WriteAttributeString(Sml.Sheet.AttrPaperSize, pageSetupSequence.Setup.iPaperSize.ToString(CultureInfo.InvariantCulture));

                if (pageSetupSequence.Setup.fUsePage)
                {
                    this._writer.WriteAttributeString(Sml.Sheet.AttrFirstPageNumber, pageSetupSequence.Setup.iPageStart.ToString(CultureInfo.InvariantCulture));
                }

                if (pageSetupSequence.Setup.fNoPls == false && pageSetupSequence.Setup.fNoOrient == false)
                {
                    // If fNoPls is 1, the value is undefined and MUST be ignored. 
                    // If fNoOrient is 1, the value is undefined and MUST be ignored.
                    this._writer.WriteAttributeString(Sml.Sheet.AttrOrientation, pageSetupSequence.Setup.fPortrait ? "portrait" : "landscape");
                }
                else
                {
                    // use landscape as default
                    this._writer.WriteAttributeString(Sml.Sheet.AttrOrientation, "landscape");
                }
                this._writer.WriteAttributeString(Sml.Sheet.AttrUseFirstPageNumber, pageSetupSequence.Setup.fUsePage ? "1" : "0");

                if (!pageSetupSequence.Setup.fNoPls)
                {
                    this._writer.WriteAttributeString(Sml.Sheet.AttrHorizontalDpi, pageSetupSequence.Setup.iRes.ToString(CultureInfo.InvariantCulture));
                    this._writer.WriteAttributeString(Sml.Sheet.AttrVerticalDpi, pageSetupSequence.Setup.iVRes.ToString(CultureInfo.InvariantCulture));
                }

                this._writer.WriteEndElement();
            }

        }

        #endregion
    }
}
