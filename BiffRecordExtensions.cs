using System;
using System.IO;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.StructuredStorage.Reader;

namespace Macrome
{
    public static class BiffRecordExtensions
    {
        public static T AsRecordType<T>(this BiffRecord record) where T : BiffRecord
        {
            byte[] recordBytes = record.GetBytes();
            using (MemoryStream ms = new MemoryStream(recordBytes))
            {
                VirtualStreamReader vsr = new VirtualStreamReader(ms);
                RecordType id = (RecordType)vsr.ReadUInt16();
                ushort len = vsr.ReadUInt16();
                var typeConstructor = typeof(T).GetConstructor(new Type[]
                    {typeof(IStreamReader), typeof(RecordType), typeof(ushort)});

                if (typeConstructor == null)
                {
                    throw new ArgumentException(string.Format("Could not find appropriate constructor for type {0}", typeof(T).FullName));
                }

                return (T)typeConstructor.Invoke(new object[] {vsr, id, len});
            }
        }
    }
}
