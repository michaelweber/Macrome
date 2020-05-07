using System.Diagnostics;
using System.Text;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.TextPropsStream)]
    public class TextPropsStream : BiffRecord
    {
        public const RecordType ID = RecordType.TextPropsStream;

        /// <summary>
        /// An FrtHeader. The FrtHeader.rt field MUST be 0x08A5.
        /// </summary>
        public FrtHeader frtHeader;

        public uint dwChecksum;

        public uint cb;

        public string rgb;

        public TextPropsStream(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeader = new FrtHeader(reader);

            this.dwChecksum = reader.ReadUInt32();
            this.cb = reader.ReadUInt32();

            var rgbBytes = reader.ReadBytes((int)this.cb);
            var codepage = Encoding.GetEncoding("ISO-8859-1"); // windows-1252 not supported by platform
            this.rgb = codepage.GetString(rgbBytes);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
