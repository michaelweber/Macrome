
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies attributes of the window used to display a sheet in a workbook 
    /// and specifies the beginning of a collection of records as defined by the Chart Sheet 
    /// Substream ABNF, Macro Sheet Substream ABNF and Worksheet Substream ABNF. The collection 
    /// of records specifies the settings of a Page Layout view for a sheet, the zoom of the 
    /// current view, the position of either frozen panes or unfrozen panes, and the selected 
    /// cells within the sheet. When this record is contained in a macro sheet substream or a 
    /// worksheet substream, it has a length of 18 bytes. When this record is contained in a chart 
    /// sheet substream, it has a length of 10 bytes which are the first 10 bytes of the original 
    /// 18 bytes record and only fSelected field is used out of all fields. This record specifies 
    /// extended properties of an associated Window1 record, and that association is 
    /// specified in Window1.
    /// </summary>
    [BiffRecord(RecordType.Window2)] 
    public class Window2 : BiffRecord
    {
        public const RecordType ID = RecordType.Window2;

        /// <summary>
        /// A bit that specifies whether the window displays formulas or values. 
        /// 
        /// If the value is 1, the window displays formulas. This field is undefined and
        /// MUST be ignored if this record is contained in a chart sheet substream.
        /// </summary>
        public bool fDspFmlaRt;

        /// <summary>
        /// A bit that specifies whether the window displays gridlines.
        /// 
        ///     Value   Meaning
        ///     0       The window does not display gridlines.
        ///     1       The window displays gridlines.
        /// </summary>
        public bool fDspGridRt;

        /// <summary>
        /// A bit that specifies whether the window displays row headings and column headings. 
        /// 
        ///     Value   Meaning
        ///     0       The window does not display row headings and column headings.
        ///     1       The window displays row headings and column headings.
        ///     
        /// This field is undefined and MUST be ignored if this record is contained in a chart sheet substream.
        /// </summary>
        public bool fDspRwColRt;

        /// <summary>
        /// A bit that specifies whether the panes in the window are frozen. 
        /// 
        /// The value MUST be 0 if either the value of colLeft is 255 or the value of rwTop 
        /// is 65535. This field is undefined and MUST be ignored if this record is 
        /// contained in a chart sheet substream.
        /// </summary>
        public bool fFrozenRt;

        /// <summary>
        /// A bit that specifies whether the window displays zero values. 
        /// 
        ///     Value   Meaning
        ///     0       The window displays cells that have a value of zero as blank.
        ///     1       The window displays cells that have a value of zero as a zero.
        ///     
        /// This field is undefined and MUST be ignored if this record is contained in a chart sheet substream.
        /// </summary>
        public bool fDspZerosRt;

        /// <summary>
        /// A bit that specifies whether the gridlines of the window are drawn in the 
        /// window‘s default foreground color or in the color specified by the value of icvHdr. 
        /// This field is undefined and MUST be ignored if this record is contained in a 
        /// chart sheet substream. 
        /// 
        ///     Value   Meaning
        ///     0       Gridlines of the window are drawn in the color as specified by the value of icvHdr.
        ///     1       Gridlines of the window are drawn in the default foreground color of the window.
        /// </summary>
        public bool fDefaultHdr;

        /// <summary>
        /// A bit that specifies whether the text is displayed in right-to-left mode in the window. 
        /// 
        ///     Value   Meaning
        ///     0       The text is displayed in left-to-right mode.
        ///     1       The text is displayed in right-to-left mode.
        ///     
        /// This field is undefined and MUST be ignored if this record is contained in a chart sheet substream.
        /// </summary>
        public bool fRightToLeft;

        /// <summary>
        /// A bit that specifies whether the window displays the outline state. This field is
        /// undefined and MUST be ignored if this record is contained in a chart sheet substream.
        /// </summary>
        public bool fDspGuts;

        /// <summary>
        /// A bit that specifies whether the panes in the window are frozen without pane 
        /// splits or frozen with pane splits. If the value of fFrozenRt is 0, the value 
        /// of fFrozenNoSplit MUST be 0. 
        /// 
        ///     Value   Meaning
        ///     0       The panes in the window are frozen with pane splits
        ///     1       The panes in the window are frozen without pane splits.
        ///     
        /// This field is undefined and MUST be ignored if this record is contained in a chart sheet substream.
        /// </summary>
        public bool fFrozenNoSplit;

        /// <summary>
        /// A bit that specifies whether the sheet tab is selected.
        /// </summary>
        public bool fSelected;

        /// <summary>
        /// A bit that specifies whether the sheet is currently being displayed in the window. 
        /// This field is undefined and MUST be ignored if this record is contained in a chart sheet substream.
        /// </summary>
        public bool fPaged;

        /// <summary>
        /// A bit that specifies whether the sheet is in Page Break Preview view. This field is 
        /// undefined and MUST be ignored if this record is contained in a chart sheet substream.
        /// </summary>
        public bool fSLV;

        /// <summary>
        /// A RwU that specifies a zero-based row index of the first visible row of the sheet. 
        /// This field is undefined and MUST be ignored if this record is contained in a chart sheet substream.
        /// </summary>
        public ushort rwTop;

        /// <summary>
        /// A ColU that specifies a zero-based column index of the logical left-most visible column. 
        /// This field is undefined and MUST be ignored if this record is contained in a chart sheet substream.
        /// </summary>
        public ushort colLeft;

        /// <summary>
        /// An Icv that specifies the color of the gridlines. MUST be less than or equal to 64. 
        /// MUST be 64 if and only if the value of fDefaultHdr is 1. This field is undefined and 
        /// MUST be ignored if this record is contained in a chart sheet substream.
        /// </summary>
        public ushort icvHdr;

        /// <summary>
        /// An unsigned integer that specifies the zoom level in the Page Break Preview view. 
        /// If the value of fSLV is 1 and this record has an associated Scl as specified in the 
        /// ABNF in Common Productions, the value of wScaleSLV is undefined and MUST be ignored. 
        /// MUST <130> be either 0 or greater than or equal to 10 and less than or equal to 400. 
        /// A value of 0 specifies the default zoom level. This field MUST NOT exist if this 
        /// record is contained in a chart sheet substream.
        /// </summary>
        public ushort wScaleSLV;

        /// <summary> An unsigned integer that specifies the zoom level in the Normal view. If the value of fSLV is 0
        /// and fPageLayoutView field of the PLV as specified in the ABNF in Common Productions is 0 and this record 
        /// has an associated Scl, then the value of wScaleNormal is undefined and MUST be ignored. MUST <131> be 
        /// either 0 or greater than or equal to 10 and less than or equal to 400. A value of 0 specifies the default 
        /// zoom level. This field MUST NOT exist if this record is contained in a chart sheet substream. </summary>
        public ushort wScaleNormal;

        public Window2(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            bool inChartSubstream = (length == 10);

            // initialize class members from stream
            ushort flags = reader.ReadUInt16();
            this.fDspFmlaRt = Utils.BitmaskToBool(flags, 0x0001);
            this.fDspGridRt = Utils.BitmaskToBool(flags, 0x0002);
            this.fDspRwColRt = Utils.BitmaskToBool(flags, 0x0004);
            this.fFrozenRt = Utils.BitmaskToBool(flags, 0x0008);
            this.fDspZerosRt = Utils.BitmaskToBool(flags, 0x0010);
            this.fDefaultHdr = Utils.BitmaskToBool(flags, 0x0020);
            this.fRightToLeft = Utils.BitmaskToBool(flags, 0x0040);
            this.fDspGuts = Utils.BitmaskToBool(flags, 0x0080);
            this.fFrozenNoSplit = Utils.BitmaskToBool(flags, 0x0100);
            this.fSelected = Utils.BitmaskToBool(flags, 0x0200);
            this.fPaged = Utils.BitmaskToBool(flags, 0x0400);
            this.fSLV = Utils.BitmaskToBool(flags, 0x0800);

            // TODO: find a generic solution for the problem that this record is different depending on the surrounding stream

            if (!inChartSubstream)
            {
                // This field is undefined and MUST be ignored if this record is contained in a chart sheet substream.
                this.rwTop = reader.ReadUInt16();

                // This field is undefined and MUST be ignored if this record is contained in a chart sheet substream.
                this.colLeft = reader.ReadUInt16();
            }


            this.icvHdr = reader.ReadUInt16();
            reader.ReadUInt16();
            this.wScaleSLV = reader.ReadUInt16();
            this.wScaleNormal = reader.ReadUInt16();

            if (!inChartSubstream)
            {
                // These fields are undefined and MUST be ignored if this record is contained in a chart sheet substream.
                reader.ReadBytes(4);
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
