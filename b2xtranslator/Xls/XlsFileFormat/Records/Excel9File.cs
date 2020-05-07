
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// EXCEL9FILE: Excel 9 File (1C0h)
    /// 
    /// The  EXCEL9FILE record indicates the file was written by Excel 2000. 
    /// It has no record data field and is C0010000h. Any application other 
    /// than Excel 2000 that edits the file should not write out this record.
    /// </summary>
    [BiffRecord(RecordType.Excel9File)] 
    public class Excel9File : BiffRecord
    {
        public const RecordType ID = RecordType.Excel9File;

        public Excel9File(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Length == 0); 
        }
    }
}
