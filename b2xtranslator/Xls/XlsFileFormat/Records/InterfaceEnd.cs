
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// INTERFACEEND: End of User Interface Records (E2h)
    /// 
    /// This record marks the end of the user interface section of the Book stream. It has no record data field.
    /// </summary>
    [BiffRecord(RecordType.InterfaceEnd)] 
    public class InterfaceEnd : BiffRecord
    {
        public const RecordType ID = RecordType.InterfaceEnd;

        public InterfaceEnd(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Length == 0); 
        }
    }
}
