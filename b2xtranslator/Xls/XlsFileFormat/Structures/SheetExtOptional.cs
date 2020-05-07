

using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// This structure specifies sheet specific data including sheet tab color and the published state of this sheet.
    /// </summary>
    public class SheetExtOptional
    {
        /// <summary>
        /// An unsigned integer that specifies the tab color of this sheet. If the tab has a color 
        /// assigned to it, the value of this field MUST be greater than or equal to 0x08 and less 
        /// than or equal to 0x3F, as specified in the color table for Icv. If this value does not 
        /// equal to icvPlain of the associated SheetExt, the value of icvPlain takes precedence. 
        /// If the tab has no color assigned to it, the value of this field MUST be 0x7F, and MUST be ignored.
        /// </summary>
        public uint icvPlain;

        /// <summary>
        /// A bit that specifies whether conditional formatting formulas are evaluated. 
        /// 
        /// MUST be one of the following: 
        /// 
        ///     Value   Meaning
        ///     0       Conditional formatting formulas in this workbook are not evaluated.
        ///     1       Conditional formatting formulas in this workbook are evaluated.
        /// </summary>
        public bool fCondFmtCalc;

        /// <summary>
        /// A bit that specifies whether this sheet is published. 
        /// 
        /// MUST be ignored when this sheet is a chart sheet, dialog sheet, or macro sheet. 
        /// MUST be a value from the following table: 
        /// 
        ///     Value   Meaning
        ///     0       The sheet is published.
        ///     1       The sheet is not published.
        /// </summary>
        public bool fNotPublished;

        // TODO: implement CFColor structure
        public byte[] color;

        public SheetExtOptional(IStreamReader reader)
        {
            uint field = reader.ReadUInt32();
            this.icvPlain = Utils.BitmaskToUInt32(field, 0x003F);
            this.fCondFmtCalc = Utils.BitmaskToBool(field, 0x0040);
            this.fNotPublished = Utils.BitmaskToBool(field, 0x0080);

            this.color = reader.ReadBytes(16);
        }
    }
}
