using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.xls.XlsFileFormat.Records
{
    [BiffRecord(RecordType.Intl)]
    public class Intl : BiffRecord
    {
        public const RecordType ID = RecordType.Intl;

        public ushort reserved;

        public Intl(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.reserved = reader.ReadUInt16();

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }

        public Intl() : base (RecordType.Intl, 2)
        {
            reserved = 0;
        }

        public override byte[] GetBytes()
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);
            bw.Write(base.GetHeaderBytes());
            bw.Write(reserved);
            return bw.GetBytesWritten();
        }
    }
}
