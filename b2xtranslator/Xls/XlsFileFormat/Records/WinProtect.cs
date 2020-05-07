
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// WINDOWPROTECT: Windows Are Protected (19h)
    /// 
    /// The WINDOWPROTECT record stores an option from the Protect Workbook dialog box.
    /// </summary>
    [BiffRecord(RecordType.WinProtect)] 
    public class WinProtect : BiffRecord
    {
        public const RecordType ID = RecordType.WinProtect;

        /// <summary>
        ///  =1 if the workbook windows are protected
        /// </summary>
        public ushort fLockWn;
        
        public WinProtect(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fLockWn = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
