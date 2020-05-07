
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// PROT4REVPASS: Shared Workbook Protection Password (1BCh)
    /// 
    /// The PROT4REV record stores an encrypted password for shared-workbook protection.
    /// </summary>
    [BiffRecord(RecordType.Prot4RevPass)] 
    public class Prot4RevPass : BiffRecord
    {
        public const RecordType ID = RecordType.Prot4RevPass;

        /// <summary>
        /// Encrypted password (if this field is 0 (zero), there is no 
        /// Shared Workbook Protection Password; the password is entered 
        /// in the Protect Shared Workbook dialog box)
        /// </summary>
        public ushort wRevPass;

        public Prot4RevPass(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.wRevPass = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
