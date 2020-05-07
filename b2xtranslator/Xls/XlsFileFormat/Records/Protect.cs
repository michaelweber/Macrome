
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// PROTECT: Protection Flag (12h)
    /// 
    /// The PROTECT record stores the protection state for a sheet or workbook.
    /// </summary>
    [BiffRecord(RecordType.Protect)] 
    public class Protect : BiffRecord
    {
        public const RecordType ID = RecordType.Protect;

        /// <summary>
        /// =1 if the sheet or workbook is protected
        /// </summary>
        public ushort fLock;

        public Protect(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fLock = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
