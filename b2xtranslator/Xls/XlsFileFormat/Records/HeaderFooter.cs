
using System;
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.HeaderFooter)] 
    public class HeaderFooter : BiffRecord
    {
        public const RecordType ID = RecordType.HeaderFooter;

        /// <summary>
        /// An FrtHeader. The frtHeader.rt field MUST be 0x089C.
        /// </summary>
        public FrtHeader frtHeader;

        /// <summary>
        /// A GUID as specified by [MS-DTYP] that specifies the current sheet view. <br/>
        /// If it is zero, it means the current sheet. <br/>
        /// Otherwise, this field MUST match the guid field of the preceding UserSViewBegin record.
        /// </summary>
        public Guid guidSView;

        /// <summary>
        /// An unsigned integer that specifies the number of characters in strHeaderEven. <br/>
        /// MUST be less than or equal to 255. <br/>
        /// MUST be zero if fHFDiffOddEven is zero.
        /// </summary>
        public ushort cchHeaderEven;

        /// <summary>
        /// An unsigned integer that specifies the number of characters in strFooterEven.<br/> 
        /// MUST be less than or equal to 255. <br/>
        /// MUST be zero if fHFDiffOddEven is zero.
        /// </summary>
        public ushort cchFooterEven;

        /// <summary>
        /// An unsigned integer that specifies the number of characters in strHeaderFirst.<br/> 
        /// MUST be less than or equal to 255. <br/>
        /// MUST be zero if fHFDiffFirst is zero.
        /// </summary>
        public ushort cchHeaderFirst;

        /// <summary>
        /// An unsigned integer that specifies the number of characters in strFooterFirst.<br/>
        /// MUST be less than or equal to 255. <br/>
        /// MUST be zero if fHFDiffFirst is zero.
        /// </summary>
        public ushort cchFooterFirst;

        /// <summary>
        /// An XLUnicodeString that specifies the header text on the even pages. <br/>
        /// The number of characters in the string MUST be equal to cchHeaderEven. <br/>
        /// The string can contain special commands, for example a placeholder for the page number, 
        /// current date or text formatting attributes. <br/>
        /// Refer to Header for more details about the string format.
        /// </summary>
        public string strHeaderEven;

        /// <summary>
        /// An XLUnicodeString that specifies the footer text on the even pages.<br/> 
        /// The number of characters in the string MUST be equal to cchFooterEven.<br/> 
        /// The string can contain special commands, for example a placeholder for the page number, 
        /// current date or text formatting attributes.<br/> 
        /// Refer to Header for more details about the string format.
        /// </summary>
        public string strFooterEven;

        /// <summary>
        /// An XLUnicodeString that specifies the header text on the first page. <br/> 
        /// The number of characters in the string MUST be equal to cchHeaderFirst. <br/> 
        /// The string can containspecial commands, for example a placeholder for the page number, 
        /// current date or text formatting attributes. <br/> 
        /// Refer to Header for more details about the string format.
        /// </summary>
        public string strHeaderFirst;

        /// <summary>
        /// An XLUnicodeString that specifies the footer text on the first page. <br/> 
        /// The number of characters in the string MUST be equal to cchFooterFirst. <br/> 
        /// The string can contain special commands, for example a placeholder for the page number, 
        /// current date or text formatting attributes. <br/> 
        /// Refer to Header for more details about the string format.
        /// </summary>
        public string strFooterFirst;

        public HeaderFooter(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeader = new FrtHeader(reader);
            this.guidSView = new Guid(reader.ReadBytes(16));
            ushort flags = reader.ReadUInt16();
            this.cchHeaderEven = reader.ReadUInt16();
            this.cchFooterEven = reader.ReadUInt16();
            this.cchHeaderFirst = reader.ReadUInt16();
            this.cchFooterFirst = reader.ReadUInt16();

            var strHeaderEvenBytes = reader.ReadBytes(this.cchHeaderEven);
            var strFooterEvenBytes = reader.ReadBytes(this.cchFooterEven);
            var strHeaderFirstBytes = reader.ReadBytes(this.cchHeaderFirst);
            var strFooterFirstBytes = reader.ReadBytes(this.cchFooterFirst);

            //this.strHeaderEven = new XLUnicodeString(reader).Value;
            //this.strFooterEven = new XLUnicodeString(reader).Value;
            //this.strHeaderFirst = new XLUnicodeString(reader).Value;
            //this.strFooterFirst = new XLUnicodeString(reader).Value;

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
