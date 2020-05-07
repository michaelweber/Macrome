using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.CommonTranslatorLib;
using b2xtranslator.OpenXmlLib.DrawingML;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using System.Xml;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class LegendMapping : AbstractChartMapping,
          IMapping<LdSequence>
    {
        public LegendMapping(ExcelContext workbookContext, ChartContext chartContext)
            : base(workbookContext, chartContext)
        {
        }

        public void Apply(LdSequence ldSequence)
        {
            // c:legend
            this._writer.WriteStartElement(Dml.Chart.Prefix, "legend", Dml.Chart.Ns);
            {
                // c:legendPos
                if (ldSequence.CrtLayout12 != null)
                {
                    // created by Excel 2007
                    writeValueElement("legendPos", mapLegendPos(ldSequence.CrtLayout12.autolayouttype));
                }
                else
                {
                    // older Excel Versions
                    writeValueElement("legendPos", mapLegendPos(ldSequence.Pos));
                }

                // c:legendEntry

                // c:layout

                // c:overlay

                // c:spPr
                if (ldSequence.FrameSequence != null)
                {
                    ldSequence.FrameSequence.Convert(new ShapePropertiesMapping(this.WorkbookContext, this.ChartContext));
                }

                // c:txPr
                if (ldSequence.TextPropsSequence != null)
                {
                    // created by Excel 2007
                    var xmlTextProps = new XmlDocument();
                    if (ldSequence.TextPropsSequence.TextPropsStream != null)
                    {
                        if (ldSequence.TextPropsSequence.TextPropsStream.rgb != null &&
                            ldSequence.TextPropsSequence.TextPropsStream.rgb != "")
                        {
                            xmlTextProps.LoadXml(ldSequence.TextPropsSequence.TextPropsStream.rgb);
                        }
                    }
                    else if (ldSequence.TextPropsSequence.RichTextStream != null)
                    {
                        if (ldSequence.TextPropsSequence.RichTextStream.rgb != null &&
                            ldSequence.TextPropsSequence.RichTextStream.rgb != "")
                        {
                            xmlTextProps.LoadXml(ldSequence.TextPropsSequence.RichTextStream.rgb);
                        }
                    }

                    // NOTE: Don't use WriteTo on the document root because it might try to 
                    // add an XML declaration to the writer (BANG!). 
                    // Use it on the top-most element node instead.
                    //
                    if (xmlTextProps.DocumentElement != null)
                    {
                        xmlTextProps.DocumentElement.WriteTo(this._writer);
                    }
                }
                else
                {
                    // TODO: older Excel Versions
                    ldSequence.AttachedLabelSequence.Convert(new TextBodyMapping(this.WorkbookContext, this.ChartContext));
                }
            }
            this._writer.WriteEndElement(); // c:legend
        }

        private string mapLegendPos(Pos pos)
        {
            string result = "r";
            if (pos.mdTopLt == Pos.PositionMode.MDCHART && pos.mdBotRt == Pos.PositionMode.MDPARENT)
            {
                // use x1 and y1 values
                if (pos.x1 < 50)
                {
                    // positioned less than 50 from the left border
                    result = "l";
                }
                else
                {
                    if (pos.y1 < 50)
                    {
                        // positioned less than 50 from the top border
                        result = "t";
                    }
                    else
                    {
                        if (pos.y1 > pos.x1)
                        {
                            // positioned at the bottom border
                            result = "b";
                        }
                    }
                }
            }
            return result;
        }

        private string mapLegendPos(CrtLayout12.AutoLayoutType autoLayoutType)
        {
            switch (autoLayoutType)
            {
                case CrtLayout12.AutoLayoutType.Bottom:
                    return "b";
                case CrtLayout12.AutoLayoutType.TopRightCorner:
                    return "tr";
                case CrtLayout12.AutoLayoutType.Top:
                    return "t";
                case CrtLayout12.AutoLayoutType.Right:
                    return "r";
                case CrtLayout12.AutoLayoutType.Left:
                    return "l";
                default:
                    return "r";
            }
        }
    }
}
