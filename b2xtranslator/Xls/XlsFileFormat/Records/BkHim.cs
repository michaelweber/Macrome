using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies image data for a sheet background.
    /// </summary>
    [BiffRecord(RecordType.BkHim)]
    public class BkHim : BiffRecord
    {
        public enum ImageFormat
        {
            Bitmap = 0x0009,
            Native = 0x000E
        }

        /// <summary>
        /// Specifies the image format.
        /// </summary>
        public ImageFormat cf;

        /// <summary>
        /// A signed integer that specifies the size of imageBlob in bytes. <br/>
        /// MUST be greater than or equal to 1.
        /// </summary>
        public int lcb;

        /// <summary>
        /// An array of bytes that specifies the image data for the given format.
        /// </summary>
        public byte[] imageBlob;

        public BkHim(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            this.cf = (ImageFormat)reader.ReadInt16();
            reader.ReadBytes(2); // skip 2 bytes
            this.lcb = reader.ReadInt32();
            this.imageBlob = reader.ReadBytes(this.lcb);
        }
    }
}
