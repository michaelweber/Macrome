

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.OpenXmlLib.SpreadsheetML;
using b2xtranslator.OpenXmlLib;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class ChartsheetMapping : AbstractOpenXmlMapping,
          IMapping<ChartSheetSequence>
    {
        ExcelContext _xlsContext;
        ChartsheetPart _chartsheetPart;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public ChartsheetMapping(ExcelContext xlsContext, ChartsheetPart targetPart)
            : base(targetPart.XmlWriter)
        {
            this._xlsContext = xlsContext;
            this._chartsheetPart = targetPart;
        }

        /// <summary>
        /// The overload apply method 
        /// Creates the Worksheet xml document 
        /// </summary>
        /// <param name="bsd">WorkSheetData</param>
        public void Apply(ChartSheetSequence chartSheetSequence)
        {
            this._writer.WriteStartDocument();
            // chartsheet
            this._writer.WriteStartElement(Sml.Sheet.ElChartsheet, Sml.Ns);
            this._writer.WriteAttributeString("xmlns", Sml.Ns);
            this._writer.WriteAttributeString("xmlns", "r", "", OpenXmlNamespaces.Relationships);
            
            var chartSheetContentSequence = chartSheetSequence.ChartSheetContentSequence;

            // sheetPr
            this._writer.WriteStartElement(Sml.Sheet.ElSheetPr, Sml.Ns);
            if (chartSheetContentSequence.CodeName != null)
            {
                // code name
                this._writer.WriteAttributeString(Sml.Sheet.AttrCodeName, chartSheetContentSequence.CodeName.codeName.Value);
            }
            // TODO: map SheetExtOptional to published and tab color

            this._writer.WriteEndElement();

            
            // sheetViews
            if (chartSheetContentSequence.WindowSequences.Count > 0)
            {
                this._writer.WriteStartElement(Sml.Sheet.ElSheetViews, Sml.Ns);

                // Note: There is a Window2 record for each Window1 record in the beginning of the workbook.
                // The index in the list corresponds to the 0-based workbookViewId attribute.
                //
                for (int window1Id = 0; window1Id < chartSheetContentSequence.WindowSequences.Count; window1Id++)
                {
                    chartSheetContentSequence.WindowSequences[window1Id].Convert(new WindowMapping(this._xlsContext, this._chartsheetPart, window1Id, true));
                }
                this._writer.WriteEndElement();
            }

            // page setup
            chartSheetContentSequence.PageSetupSequence.Convert(new PageSetupMapping(this._xlsContext, this._chartsheetPart));

            // header and footer
            // TODO: header and footer

            // drawing
            this._writer.WriteStartElement(Sml.Sheet.ElDrawing, Sml.Ns);
            this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, this._chartsheetPart.DrawingsPart.RelIdToString);
            chartSheetContentSequence.Convert(new DrawingMapping(this._xlsContext, this._chartsheetPart.DrawingsPart, true));

            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }
    }
}
