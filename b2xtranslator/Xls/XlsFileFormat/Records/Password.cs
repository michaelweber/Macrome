
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// PASSWORD: Protection Password (13h)
    /// 
    /// The PASSWORD record contains the encrypted password for a protected sheet or workbook.  
    /// 
    /// Note: this record specifies a sheet-level or workbook-level protection password, 
    /// as opposed to the FILEPASS record, which specifies a file password.
    /// </summary>
    [BiffRecord(RecordType.Password)] 
    public class Password : BiffRecord
    {
        public const RecordType ID = RecordType.Password;

        /// <summary>
        /// Encrypted password
        /// </summary>
        public ushort wPassword;

        public Password(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.wPassword = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
