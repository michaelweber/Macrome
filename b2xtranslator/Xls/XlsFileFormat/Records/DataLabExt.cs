

using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the beginning of a collection of records as defined by 
    /// the Chart Sheet Substream ABNF. The collection specifies an extended data label.
    /// </summary>
    [BiffRecord(RecordType.DataLabExt)]
    public class DataLabExt : BiffRecord
    {
        public const RecordType ID = RecordType.DataLabExt;

        public FrtHeader frtHeader;

        public DataLabExt(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeader = new FrtHeader(reader);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
