

using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.OpenXmlLib.DrawingML;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class DataLabelMapping : AbstractChartMapping,
          IMapping<ChartFormatsSequence>
    {
        SeriesFormatSequence _seriesFormatSequence;

        public DataLabelMapping(ExcelContext workbookContext, ChartContext chartContext, SeriesFormatSequence parentSeriesFormatSequence)
            : base(workbookContext, chartContext)
        {
            this._seriesFormatSequence = parentSeriesFormatSequence;
        }

        public void Apply(ChartFormatsSequence chartFormatSequence)
        {
            if (chartFormatSequence.AttachedLabelSequences.Count != 0)
            {
                var attachedLabelSequence = chartFormatSequence.AttachedLabelSequences[0];

                this._writer.WriteStartElement(Dml.Chart.Prefix, Dml.Chart.ElDLbls, Dml.Chart.Ns);
                {
                    //<xsd:element name="dLbl" type="CT_DLbl" minOccurs="0" maxOccurs="unbounded"/>
                    // TODO

                    //<xsd:element name="dLblPos" type="CT_DLblPos" minOccurs="0" maxOccurs="1">
                    // TODO

                    //<xsd:element name="leaderLines" type="CT_ChartLines" minOccurs="0" maxOccurs="1">
                    // TODO

                    //<xsd:element name="numFmt" type="CT_NumFmt" minOccurs="0" maxOccurs="1">
                    // TODO

                    //<xsd:element name="showLeaderLines" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                    // TODO

                    if (attachedLabelSequence != null)
                    {
                        if (attachedLabelSequence.Text != null)
                        {
                            //<xsd:element name="showLegendKey" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                            if (attachedLabelSequence.Text.fShowKey)
                            {
                                writeValueElement("showLegendKey", "1");
                            }
                        }

                        if (attachedLabelSequence.DataLabExtContents != null)
                        {
                            //<xsd:element name="separator" type="xsd:string" minOccurs="0" maxOccurs="1">
                            this._writer.WriteElementString(Dml.Chart.Prefix, "separator", Dml.Chart.Ns, attachedLabelSequence.DataLabExtContents.rgchSep.Value);

                            //<xsd:element name="showBubbleSize" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                            if (attachedLabelSequence.DataLabExtContents.fBubSizes)
                            {
                                writeValueElement("showBubbleSize", "1");
                            }

                            //<xsd:element name="showCatName" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                            if (attachedLabelSequence.DataLabExtContents.fCatName)
                            {
                                writeValueElement("showCatName", "1");
                            }

                            //<xsd:element name="showPercent" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                            if (attachedLabelSequence.DataLabExtContents.fPercent)
                            {
                                writeValueElement("showPercent", "1");
                            }

                            //<xsd:element name="showSerName" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                            if (attachedLabelSequence.DataLabExtContents.fSerName)
                            {
                                writeValueElement("showSerName", "1");
                            }

                            //<xsd:element name="showVal" type="CT_Boolean" minOccurs="0" maxOccurs="1">
                            if (attachedLabelSequence.DataLabExtContents.fValue)
                            {
                                writeValueElement("showVal", "1");
                            }
                        }

                    }
                    //<xsd:element name="spPr" type="a:CT_ShapeProperties" minOccurs="0" maxOccurs="1">
                    if (attachedLabelSequence.FrameSequence != null)
                    {
                        attachedLabelSequence.FrameSequence.Convert(new ShapePropertiesMapping(this.WorkbookContext, this.ChartContext));
                    }

                    //<xsd:element name="txPr" type="a:CT_TextBody" minOccurs="0" maxOccurs="1">
                    // TODO

                    //<xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
                    // TODO

                }
                this._writer.WriteEndElement();
            }

        }
    }
}
