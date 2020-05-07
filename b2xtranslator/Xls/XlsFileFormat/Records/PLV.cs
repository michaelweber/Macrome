
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the settings of a Page Layout view for a sheet.
    /// </summary>
    [BiffRecord(RecordType.PLV)] 
    public class PLV : BiffRecord
    {
        public const RecordType ID = RecordType.PLV;

        /// <summary>
        /// An FrtHeader. The frtHeader.rt field MUST be 0x088B.
        /// </summary>
        public FrtHeader frtHeader;

        /// <summary>
        /// An unsigned integer that specifies zoom scale as a percentage for the 
        /// Page Layout view of the current sheet. For example, if the value is 107, 
        /// then the zoom scale is 107%. The value 0 means that the zoom scale has not 
        /// been set. If the value is nonzero, it MUST be greater than or equal 
        /// to 10 and less than or equal to 400.
        /// </summary>
        public ushort wScalePLV;

        /// <summary>
        /// A bit that specifies whether the sheet is in the Page Layout view. 
        /// If the fSLV in Window2 record is 1 for this sheet, it MUST be 0
        /// </summary>
        public bool fPageLayoutView;

        /// <summary>
        /// A bit that specifies whether the application displays the ruler.
        /// </summary>
        public bool fRulerVisible;

        /// <summary>
        /// A bit that specifies whether the margins between pages are hidden in the Page Layout view.
        /// </summary>
        public bool fWhitespaceHidden;


        public PLV(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeader = new FrtHeader(reader);
            this.wScalePLV = reader.ReadUInt16();

            ushort flags = reader.ReadUInt16();
            this.fPageLayoutView = Utils.BitmaskToBool(flags, 0x0001);
            this.fRulerVisible = Utils.BitmaskToBool(flags, 0x0002);
            this.fWhitespaceHidden = Utils.BitmaskToBool(flags, 0x0004);
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
