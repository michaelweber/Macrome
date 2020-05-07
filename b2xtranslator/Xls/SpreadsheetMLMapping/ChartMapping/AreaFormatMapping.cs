namespace b2xtranslator.SpreadsheetMLMapping
{
    // TODO: no longer needed
    //public class AreaFormatMapping : AbstractChartMapping,
    //      IMapping<AreaFormat>
    //{
    //    GelFrameSequence _gelFrameSequence;

    //    public AreaFormatMapping(ExcelContext workbookContext, ChartContext chartContext, GelFrameSequence gelFrameSequence)
    //        : base(workbookContext, chartContext)
    //    {
    //         _gelFrameSequence = gelFrameSequence;
    //    }

    //    public void Apply(AreaFormat areaFormat)
    //    {
    //        if (areaFormat.fls == 1)
    //        {
    //            // SOLID FILL

    //            RGBColor fillColor;
    //            if (this.ChartSheetContentSequence.Palette != null && areaFormat.icvFore >= 0x0000 && areaFormat.icvFore <= 0x0041)
    //            {
    //                // there is a valid palette color set
    //                fillColor = this.ChartSheetContentSequence.Palette.rgbColorList[areaFormat.icvFore];
    //            }
    //            else
    //            {
    //                fillColor = areaFormat.rgbFore;
    //            }
    //            if(fillColor != null)
    //            {
    //                // <a:solidFill>
    //                _writer.WriteStartElement(Dml.Prefix, "solidFill", Dml.Ns);
    //                {
    //                    // <a:srgbClr val="000000"/>
    //                    writeValueElement(Dml.Prefix, "srgbClr", Dml.Ns, fillColor.SixDigitHexCode);
    //                }
    //                _writer.WriteEndElement();
    //            }
    //        }
    //    }
    //}
}
