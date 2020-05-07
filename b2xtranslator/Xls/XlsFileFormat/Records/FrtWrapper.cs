

using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record wraps around a non-Future Record Type (FRT) record and converts it into an FRT record.
    /// </summary>
    [BiffRecord(RecordType.FrtWrapper)]
    public class FrtWrapper : BiffRecord
    {
        public const RecordType ID = RecordType.FrtWrapper;

        public FrtHeaderOld frtHeaderOld;

        public BiffRecord wrappedRecord;

        public FrtWrapper(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            this.frtHeaderOld = new FrtHeaderOld(reader);

            this.wrappedRecord = BiffRecord.ReadRecord(reader);

            // skip padding bytes
            this.Reader.BaseStream.Position = this.Offset + this.Length;
        }
    }
}
