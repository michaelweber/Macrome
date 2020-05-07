
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// MMS: ADDMENU/DELMENU Record Group Count (C1h)
    /// 
    /// This record stores the number of ADDMENU  groups and DELMENU  groups in the Book  stream.
    /// </summary>
    [BiffRecord(RecordType.Mms)] 
    public class Mms : BiffRecord
    {
        public const RecordType ID = RecordType.Mms;

        /// <summary>
        /// Number of ADDMENU record groups
        /// </summary>
        public byte caitm;

        /// <summary>
        /// Number of DELMENU record groups
        /// </summary>
        public byte cditm;  

        public Mms(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.caitm = reader.ReadByte();
            this.cditm = reader.ReadByte();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
