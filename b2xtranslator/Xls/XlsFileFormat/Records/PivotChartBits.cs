
using System;
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.PivotChartBits)]
    public class PivotChartBits : BiffRecord
    {
        public const RecordType ID = RecordType.PivotChartBits;

        public ushort rt;

        public bool fGXHide;

        public PivotChartBits(IStreamReader reader, RecordType id, ushort length)
            :base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.rt = reader.ReadUInt16();

            //unused
            reader.ReadBytes(2);

            var bytes = reader.ReadBytes(2);

            this.fGXHide = BitConverter.ToBoolean(bytes, 0x1);

        }
    }
}
