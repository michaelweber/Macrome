
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the footer text of the current sheet when printed.
    /// </summary>
    [BiffRecord(RecordType.Footer)] 
    public class Footer : BiffRecord
    {
        public const RecordType ID = RecordType.Footer;

        /// <summary>
        /// An XLUnicodeString that specifies the footer text for the current sheet. 
        /// It is optional and exists only if the record size is not zero. The footer 
        /// text appears at the bottom of every page when printed. The length of the 
        /// text MUST be less than or equal to 255. The footer text can contain special 
        /// commands, for example a placeholder for the page number, current date or text 
        /// formatting attributes, as specified in the ABNF grammar for special
        /// commands as specified in Header.
        /// </summary>
        public XLUnicodeString footerText; 

        public Footer(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            if (this.Length > 0)
            {
                this.footerText = new XLUnicodeString(reader);
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
