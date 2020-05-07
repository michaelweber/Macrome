

using System;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.SpreadsheetML;
using b2xtranslator.Spreadsheet.XlsFileFormat;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class WindowMapping : AbstractOpenXmlMapping,
          IMapping<WindowSequence>
    {
        ExcelContext _xlsContext;
        ChartsheetPart _chartsheetPart;
        int _window1Id = 0;
        bool _isChartsheetSubstream;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="xlsContext">The excel context object</param>
        public WindowMapping(ExcelContext xlsContext, ChartsheetPart chartsheetPart, int window1Id, bool isChartsheetSubstream)
            : base(chartsheetPart.XmlWriter)
        {
            this._xlsContext = xlsContext;
            this._chartsheetPart = chartsheetPart;
            this._window1Id = window1Id;
            this._isChartsheetSubstream = isChartsheetSubstream;
        }

        #region IMapping<WindowSequence> Members

        public void Apply(WindowSequence windowSequence)
        {
            this._writer.WriteStartElement(Sml.Sheet.ElSheetView, Sml.Ns);

            this._writer.WriteAttributeString(Sml.Sheet.AttrTabSelected, windowSequence.Window2.fSelected ? "1" : "0");

            // zoomScale
            if (windowSequence.Scl != null)
            {
                // zoomScale is a percentage in the range 10-400
                int zoomScale = Math.Min(400, Math.Max(10, windowSequence.Scl.nscl * 100 / windowSequence.Scl.dscl));
                this._writer.WriteAttributeString(Sml.Sheet.AttrZoomScale, zoomScale.ToString());
            }

            this._writer.WriteAttributeString(Sml.Sheet.AttrWorkbookViewId, this._window1Id.ToString());
            // TODO: complete mapping, certain fields must be ignored when in chart sheet substream

            this._writer.WriteEndElement();
        }

        #endregion
    }
}
