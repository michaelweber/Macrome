
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// Window1: Window Information (3Dh)
    /// 
    /// The Window1 record contains workbook-level window attributes. 
    /// 
    /// The xWn and yWn fields contain the location of the window in units of 1/20th of a point, 
    /// relative to the upper-left corner of the Excel window client area. 
    /// 
    /// The dxWn and dyWn fields contain the window size, also in units of 1/20th  of a point.
    /// </summary>
    [BiffRecord(RecordType.Window1)] 
    public class Window1 : BiffRecord
    {
        public const RecordType ID = RecordType.Window1;

        /// <summary>
        /// Horizontal position of the window in units of 1/20th of a point.
        /// </summary>
        public ushort xWn;
	
        /// <summary>
        /// Vertical position of the window in units of 1/20th of a point.
        /// </summary>
        public ushort yWn;

        /// <summary>
        /// Width of the window in units of 1/20th of a point.
        /// </summary>
        public ushort dxWn;
	
        /// <summary>
        /// Height of the window in units of 1/20th of a point.
        /// </summary>
        public ushort dyWn;
	
        /// <summary>
        /// Option flags
        /// </summary>
        public ushort grbit;
	
    	/// <summary>
    	/// Index of the selected workbook tab (0-based).
    	/// </summary>
        public ushort itabCur;
	
        /// <summary>
        /// Index of the first displayed workbook tab (0-based).
        /// </summary>
        public ushort itabFirst;
	
        /// <summary>
        /// Number of workbook tabs that are selected.
        /// </summary>
        public ushort ctabSel;

        /// <summary>
        /// Ratio of the width of the workbook tabs to the width of the horizontal scroll bar; 
        /// to obtain the ratio, convert to decimal and then divide by 1000.
        /// </summary>
        public ushort wTabRatio;	

        // The grbit field contains the following option flags:
        // Field                        Offset	Bits    Mask	Name	Contents
        public bool fHidden;        //  0	    0       01h		=1 if the window is hidden
        public bool fIconic; 	    //     	    1	    02h     =1 if the window is currently displayed as an icon
        public bool reserved0;      //     	    2	    04h	
        public bool fDspHScroll; 	//     	    3	    08h	    =1 if the horizontal scroll bar is displayed
        public bool fDspVScroll;    //    	    4	    10h	    =1 if the vertical scroll bar is displayed
        public bool fBotAdornment;  //     	    5	    20h	    =1 if the workbook tabs are displayed
        public bool fNoAFDateGroup; //     	    6	    40h	    =1 if the AutoFilter should not group dates (Excel 11 (2003) behavior), new for Office Excel 2007
        public bool reserved1;      //     	    7	    80h
        public byte reserved2;      //  1       7–0	    FFh	

        public Window1(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.xWn = reader.ReadUInt16();
            this.yWn = reader.ReadUInt16();
            this.dxWn = reader.ReadUInt16();
            this.dyWn = reader.ReadUInt16();

            this.grbit = reader.ReadUInt16();

            this.fHidden = Utils.BitmaskToBool(this.grbit, 0x01);
            this.fIconic = Utils.BitmaskToBool(this.grbit, 0x02);
            this.reserved0 = Utils.BitmaskToBool(this.grbit, 0x04);
            this.fDspHScroll = Utils.BitmaskToBool(this.grbit, 0x08);
            this.fDspVScroll = Utils.BitmaskToBool(this.grbit, 0x10);
            this.fBotAdornment = Utils.BitmaskToBool(this.grbit, 0x20);
            this.fNoAFDateGroup = Utils.BitmaskToBool(this.grbit, 0x40);
            this.reserved1 = Utils.BitmaskToBool(this.grbit, 0x80);
            this.reserved2 = Utils.BitmaskToByte(this.grbit, 0xFF00);

            this.itabCur = reader.ReadUInt16();
            this.itabFirst = reader.ReadUInt16();
            this.ctabSel = reader.ReadUInt16();
            this.wTabRatio = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
