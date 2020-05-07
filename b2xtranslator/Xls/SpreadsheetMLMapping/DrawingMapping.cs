

using System;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OfficeDrawing;
using b2xtranslator.OpenXmlLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.Tools;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class DrawingMapping : AbstractOpenXmlMapping,
        IMapping<ChartSheetContentSequence>,
        IMapping<ObjectsSequence>
    {
        ExcelContext _xlsContext;
        DrawingsPart _drawingsPart;

        bool _isChartsheet;

        public DrawingMapping(ExcelContext xlsContext, DrawingsPart targetPart, bool isChartsheet)
            : base(targetPart.XmlWriter)
        {
            this._xlsContext = xlsContext;
            this._drawingsPart = targetPart;

            this._isChartsheet = isChartsheet;
        }

        #region IMapping<ChartSheetContentSequence> Members

        public void Apply(ChartSheetContentSequence chartSheetContentSequence)
        {
            this._writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElWsDr, Dml.SpreadsheetDrawing.Ns);
            this._writer.WriteAttributeString("xmlns", Dml.SpreadsheetDrawing.Prefix, "", Dml.SpreadsheetDrawing.Ns);
            this._writer.WriteAttributeString("xmlns", Dml.Prefix, "", Dml.Ns);

            if (this._isChartsheet)
            {
                this._writer.WriteStartElement(Dml.SpreadsheetDrawing.ElAbsoluteAnchor, Dml.SpreadsheetDrawing.Ns);
                {
                    var chart = chartSheetContentSequence.ChartFormatsSequence.Chart;

                    // NOTE: Excel seems to somehow round the pos and ext values. The exact calculation is not documented.
                    //   Besides, Excel might write negative values which are corrected to 0 by Excel on load time.
                    //
                    // xdr:pos
                    this._writer.WriteStartElement(Dml.SpreadsheetDrawing.ElPos, Dml.SpreadsheetDrawing.Ns);
                    this._writer.WriteAttributeString(Dml.BaseTypes.AttrX, Math.Max(0, new PtValue(chart.x.Value).ToEmu()).ToString());
                    this._writer.WriteAttributeString(Dml.BaseTypes.AttrY, Math.Max(0, new PtValue(chart.y.Value).ToEmu()).ToString());
                    this._writer.WriteEndElement();

                    // xdr:ext
                    this._writer.WriteStartElement(Dml.SpreadsheetDrawing.ElExt, Dml.SpreadsheetDrawing.Ns);
                    this._writer.WriteAttributeString(Dml.BaseTypes.AttrCx, Math.Max(0, new PtValue(chart.dx.Value).ToEmu()).ToString());
                    this._writer.WriteAttributeString(Dml.BaseTypes.AttrCy, Math.Max(0, new PtValue(chart.dy.Value).ToEmu()).ToString());
                    this._writer.WriteEndElement();

                    // insert EG_ObjectChoices type
                    insertObjectChoices(chartSheetContentSequence);
                }
                this._writer.WriteEndElement(); // absoluteAnchor
            }
            else
            {
                // embedded drawing
            }

            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }

        #endregion

        #region IMapping<ObjectsSequence> Members

        /// <summary>
        /// Mapping definition for embedded objects
        /// </summary>
        /// <param name="objectsSequence"></param>
        public void Apply(ObjectsSequence objectsSequence)
        {
            this._writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElWsDr, Dml.SpreadsheetDrawing.Ns);
            this._writer.WriteAttributeString("xmlns", Dml.SpreadsheetDrawing.Prefix, "", Dml.SpreadsheetDrawing.Ns);
            this._writer.WriteAttributeString("xmlns", Dml.Prefix, "", Dml.Ns);

            foreach (var drawingsGroup in objectsSequence.DrawingsGroup)
            {
                // TODO: currently only embedded charts are mapped. Shapes and images are not yet implemented.
                //    The check on the type of object would have to be removed here once shapes and images are supported.
                //
                var objGroup = drawingsGroup.Objects.Find(p => p.ChartSheetSequence != null);
                if (objGroup != null)
                {
                    var msoDrawing = drawingsGroup.MsoDrawing;

                    // find OfficeArtClientAnchorSheet with drawing
                    var container = msoDrawing.rgChildRec as RegularContainer;

                    if (container != null)
                    {
                        var clientAnchor = container.FirstDescendantWithType<ClientAnchor>();
                        if (clientAnchor != null)
                        {
                            var oartClientAnchor = new OfficeArtClientAnchorSheet(clientAnchor.RawData);

                            // xdr:twoCellAnchor
                            this._writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElTwoCellAnchor, Dml.SpreadsheetDrawing.Ns);
                            string editAs = "absolute";
                            if (oartClientAnchor.fSize && oartClientAnchor.fMove)
                            {
                                // Move and resize with anchor cells
                                editAs = "twoCell";
                            }
                            else if (!oartClientAnchor.fSize && oartClientAnchor.fMove)
                            {
                                // Move with cells but do not resize
                                editAs = "oneCell";
                            }
                            this._writer.WriteAttributeString(Dml.SpreadsheetDrawing.AttrEditAs, editAs);
                            {
                                // xdr:from
                                this._writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElFrom, Dml.SpreadsheetDrawing.Ns);
                                this._writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElCol, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.colL.ToString());
                                this._writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElColOff, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.dxL.ToString());
                                this._writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElRow, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.rwT.ToString());
                                this._writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElRowOff, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.dyT.ToString());
                                this._writer.WriteEndElement(); // xdr:from

                                // xdr:to
                                this._writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElTo, Dml.SpreadsheetDrawing.Ns);
                                this._writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElCol, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.colR.ToString());
                                this._writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElColOff, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.dxR.ToString());
                                this._writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElRow, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.rwB.ToString());
                                this._writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElRowOff, Dml.SpreadsheetDrawing.Ns, oartClientAnchor.dyB.ToString());
                                this._writer.WriteEndElement(); // xdr:to

                                var objectGroup = drawingsGroup.Objects.Find(p => p.ChartSheetSequence != null);
                                if (objectGroup != null)
                                {
                                    var chartSheetContentSequence = objectGroup.ChartSheetSequence.ChartSheetContentSequence;
                                    insertObjectChoices(chartSheetContentSequence);
                                }
                            }
                            this._writer.WriteEndElement(); // xdr:twoCellAnchor
                        }
                    }
                }
            }

            this._writer.WriteEndElement();
            this._writer.WriteEndDocument();

            this._writer.Flush();
        }

        #endregion

        private void insertObjectChoices(ChartSheetContentSequence chartSheetContentSequence)
        {
            this._writer.WriteStartElement(Dml.SpreadsheetDrawing.ElGraphicFrame, Dml.SpreadsheetDrawing.Ns);
            {
                // TODO: add graphic properties
                this._writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElNvGraphicFramePr, Dml.SpreadsheetDrawing.Ns);
                {
                    this._writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElCNvPr, Dml.SpreadsheetDrawing.Ns);
                    this._writer.WriteAttributeString(Dml.DocumentProperties.AttrId, this._drawingsPart.RelId.ToString());
                    this._writer.WriteAttributeString(Dml.DocumentProperties.AttrName, "Shape");
                    this._writer.WriteEndElement(); // xdr:cNvPr

                    this._writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElCNvGraphicFramePr, Dml.SpreadsheetDrawing.Ns);
                    this._writer.WriteStartElement(Dml.Prefix, Dml.DocumentProperties.ElGraphicFrameLocks, Dml.Ns);
                    this._writer.WriteAttributeString(Dml.DocumentProperties.AttrNoGrp, "1");
                    this._writer.WriteEndElement(); // a:graphicFrameLocks
                    this._writer.WriteEndElement(); // xdr:cNvGraphicFramePr
                }
                this._writer.WriteEndElement(); // xdr:nvGraphicFramePr

                // xdr:xfrm
                this._writer.WriteStartElement(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElXfrm, Dml.SpreadsheetDrawing.Ns);
                {
                    this._writer.WriteStartElement(Dml.Prefix, Dml.BaseTypes.ElOff, Dml.Ns);
                    this._writer.WriteAttributeString(Dml.BaseTypes.AttrX, "0");
                    this._writer.WriteAttributeString(Dml.BaseTypes.AttrY, "0");
                    this._writer.WriteEndElement(); // a:off

                    this._writer.WriteStartElement(Dml.Prefix, Dml.BaseTypes.ElExt, Dml.Ns);
                    this._writer.WriteAttributeString(Dml.BaseTypes.AttrCx, "0");
                    this._writer.WriteAttributeString(Dml.BaseTypes.AttrCy, "0");
                    this._writer.WriteEndElement(); // a:ext
                }
                this._writer.WriteEndElement(); // xdr:xfrm


                this._writer.WriteStartElement(Dml.GraphicalObject.ElGraphic, Dml.Ns);
                {
                    this._writer.WriteStartElement(Dml.GraphicalObject.ElGraphicData, Dml.Ns);
                    this._writer.WriteAttributeString(Dml.GraphicalObject.AttrUri, Dml.Chart.Ns);

                    // create and convert chart part
                    var chartPart = this._drawingsPart.AddChartPart();
                    var chartContext = new ChartContext(chartPart, chartSheetContentSequence, this._isChartsheet ? ChartContext.ChartLocation.Chartsheet : ChartContext.ChartLocation.Embedded);
                    chartSheetContentSequence.Convert(new ChartMapping(this._xlsContext, chartContext));

                    this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElChart, Dml.Chart.Ns);
                    this._writer.WriteAttributeString("r", "id", OpenXmlNamespaces.Relationships, chartPart.RelIdToString);

                    this._writer.WriteEndElement(); // c:chart

                    this._writer.WriteEndElement(); // a:graphicData
                }
                this._writer.WriteEndElement(); // a:graphic
            }
            this._writer.WriteEndElement(); // a:graphicFrame

            this._writer.WriteElementString(Dml.SpreadsheetDrawing.Prefix, Dml.SpreadsheetDrawing.ElClientData, Dml.SpreadsheetDrawing.Ns, string.Empty);
        }
    }
}
