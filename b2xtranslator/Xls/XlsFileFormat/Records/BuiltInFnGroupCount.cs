using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// FNGROUPCOUNT: Built-in Function Group Count (9Ch)
    /// 
    /// This record stores the number of built-in function groups 
    /// (Financial, Math & Trig, Date & Time, and so on) in the current version of Excel.
    /// </summary>
    [BiffRecord(RecordType.BuiltInFnGroupCount)] 
    public class BuiltInFnGroupCount : BiffRecord
    {
        public const RecordType ID = RecordType.BuiltInFnGroupCount;

        /// <summary>
        /// Number of built-in function groups
        /// </summary>
        public ushort cFnGroup;
        
        public BuiltInFnGroupCount(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.cFnGroup = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
