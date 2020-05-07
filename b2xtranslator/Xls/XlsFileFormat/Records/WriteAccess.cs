
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// WRITEACCESS: Write Access User Name (5Ch)
    /// 
    /// This record contains the user name, which is the name entered when installing Excel.
    /// </summary>
    [BiffRecord(RecordType.WriteAccess)] 
    public class WriteAccess : BiffRecord
    {
        public const RecordType ID = RecordType.WriteAccess;

        public WriteAccess(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            // TODO: place code here
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
