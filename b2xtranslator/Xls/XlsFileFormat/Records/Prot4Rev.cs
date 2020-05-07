
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// PROT4REV: Shared Workbook Protection Flag (1AFh)
    /// 
    /// The PROT4REV  record stores a shared-workbook protection flag.
    /// </summary>
    [BiffRecord(RecordType.Prot4Rev)] 
    public class Prot4Rev : BiffRecord
    {
        public const RecordType ID = RecordType.Prot4Rev;

        /// <summary>
        /// =1 if the Sharing with Track Changes option is on (Protect Shared Workbook dialog box)
        /// </summary>
        public ushort fRevLock;

        public Prot4Rev(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fRevLock = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
