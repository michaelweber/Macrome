namespace b2xtranslator.SpreadsheetMLMapping
{
    // TODO: no longer used
    //public class LineFormatMapping : AbstractChartMapping,
    //      IMapping<LineFormat>
    //{

    //    public LineFormatMapping(ExcelContext workbookContext, ChartContext chartContext)
    //        : base(workbookContext, chartContext)
    //    {

    //    }

    //    public void Apply(LineFormat lineFormat)
    //    {
    //        // it seems that Excel 2007 doesn't convert hairline lines
    //        if (lineFormat.we != LineFormat.LineWeight.Hairline)
    //        {
    //            // a:ln
    //            _writer.WriteStartElement(Dml.Prefix, "ln", Dml.Ns);
    //            {
    //                _writer.WriteAttributeString("", "w", "", mapLineWidth(lineFormat.we));

    //                //<a:solidFill>
    //                _writer.WriteStartElement(Dml.Prefix, "solidFill", Dml.Ns);
    //                {
    //                    //<a:srgbClr val="000000"/>
    //                    _writer.WriteStartElement(Dml.Prefix, "srgbClr", Dml.Ns);
    //                    {
    //                        _writer.WriteAttributeString("", "val", "", lineFormat.rgb.SixDigitHexCode);
    //                    }
    //                    _writer.WriteEndElement();
    //                }
    //                _writer.WriteEndElement();

    //                // <a:prstDash val="solid"/>
    //                writeValueElement(Dml.Prefix, "prstDash", Dml.Ns, mapLineStyle(lineFormat.lns));
    //            }
    //            _writer.WriteEndElement();



    //        }
    //    }

    //    private string mapLineWidth(LineFormat.LineWeight lineWeight)
    //    {
    //        switch (lineWeight)
    //        {
    //            case LineFormat.LineWeight.Hairline:
    //                return "0";
    //            case LineFormat.LineWeight.Narrow:
    //                return "12700";
    //            case LineFormat.LineWeight.Medium:
    //                return "25400";
    //            case LineFormat.LineWeight.Wide:
    //                return "38100";
    //        }
    //    }

    //    private string mapLineStyle(LineFormat.LineStyle lns)
    //    {
    //        switch (lns)
    //        {
    //            case LineFormat.LineStyle.Solid:
    //                return "solid";
    //            case LineFormat.LineStyle.Dash:
    //                return "dash";
    //            case LineFormat.LineStyle.Dot:
    //                return "dot";
    //            case LineFormat.LineStyle.DashDot:
    //                return "dashDot";
    //            case LineFormat.LineStyle.DashDotDot:
    //                return "sysDashDotDot";

    //            case LineFormat.LineStyle.None:
    //                return "solid"; 
    //            case LineFormat.LineStyle.DarkGrayPattern:
    //                return "solid"; 
    //            case LineFormat.LineStyle.MediumGrayPattern:
    //                return "solid"; 
    //            case LineFormat.LineStyle.LightGrayPattern:
    //                return "solid"; 
    //            default:
    //                return "solid"; 
    //        }
    //    }
    //}
}
