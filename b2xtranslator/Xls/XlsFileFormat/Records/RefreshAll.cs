
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// REFRESHALL: Refresh Flag (1B7h)
    /// 
    /// This record stores an option flag.
    /// </summary>
    [BiffRecord(RecordType.RefreshAll)] 
    public class RefreshAll : BiffRecord
    {
        public const RecordType ID = RecordType.RefreshAll;

        /// <summary>
        /// =1 then Refresh All should be done on all external data ranges and PivotTables when loading the workbook (the default is =0)
        /// </summary>
        public ushort fRefreshAll;

        public RefreshAll(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fRefreshAll = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
