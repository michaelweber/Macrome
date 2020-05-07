
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.ShapePropsStream)]
    public class ShapePropsStream : BiffRecord
    {
        public const RecordType ID = RecordType.ShapePropsStream;

        public FrtHeader frtHeader;

        /// <summary>
        /// An unsigned integer that specifies the chart element that the shape 
        /// formatting properties in this record apply to.
        /// </summary>
        public ushort wObjContext;

        /// <summary>
        /// An unsigned integer that specifies the checksum of the shape formatting properties related to this record.
        /// </summary>
        public uint dwChecksum;

        /// <summary>
        /// An unsigned integer that specifies the length of the character array in the rgb field.
        /// </summary>
        public uint cb;

        /// <summary>
        /// An array of ANSI characters, whose length is specified by cb, 
        /// that contains the XML representation of the shape formatting properties 
        /// as defined in [ECMA-376] Part 4, section 5.7.2.198.
        /// </summary>
        public byte[] rgb;

        public ShapePropsStream(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.frtHeader = new FrtHeader(reader);
            this.wObjContext = reader.ReadUInt16();
            reader.ReadBytes(2); //unused
            this.dwChecksum = reader.ReadUInt32();
            this.cb = reader.ReadUInt32();
            this.rgb = reader.ReadBytes((int)this.cb);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
