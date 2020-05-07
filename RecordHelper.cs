using System;
using System.Collections.Generic;
using System.IO;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace Macrome
{
    public class RecordHelper
    {
        public static List<BiffRecord> ParseBiffStreamBytes(byte[] bytes)
        {
            List<BiffRecord> records = new List<BiffRecord>();
            MemoryStream ms = new MemoryStream(bytes);
            VirtualStreamReader vsr = new VirtualStreamReader(ms);

            while (vsr.BaseStream.Position < vsr.BaseStream.Length)
            {
                RecordType id = (RecordType)vsr.ReadUInt16();
                UInt16 length = vsr.ReadUInt16();

                BiffRecord br = new BiffRecord(vsr, id, length);

                vsr.ReadBytes(length);
                records.Add(br);
            }

            return records;
        }

        public static byte[] ConvertBiffRecordsToBytes(List<BiffRecord> records)
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);
            foreach (var record in records)
            {
                bw.Write(record.GetBytes());
            }
            return bw.GetBytesWritten();
        }

    }
}
