

using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies the number of non-empty rows and the number of non-empty cells in the longest row of a Graph object.
    /// </summary>
    [BiffRecord(RecordType.Dimensions)]
    public class Dimensions : BiffRecord
    {
        public const RecordType ID = RecordType.Dimensions;

        /// <summary>
        /// A RwLongU that specifies the first row in the sheet that contains a used cell.
        /// </summary>
        public uint rwMic;

        /// <summary>
        /// An unsigned integer that specifies the number of non-empty cells in the 
        /// longest row in the data sheet of a Graph object. 
        /// 
        /// MUST be less than or equal to 0x00000F9F.
        /// </summary>
        public uint rwMac;

        /// <summary>
        /// A ColU that specifies the first column in the sheet that contains a used cell.
        /// </summary>
        public ushort colMic;

        /// <summary>
        /// An unsigned integer that specifies the number of non-empty rows in the 
        /// data sheet of a Graph object. 
        /// 
        /// MUST be less than or equal to 0x00FF.
        /// </summary>
        public ushort colMac;

        public Dimensions(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.rwMic = reader.ReadUInt32();
            this.rwMac = reader.ReadUInt32();
            this.colMic = reader.ReadUInt16();
            this.colMac = reader.ReadUInt16();
            reader.ReadBytes(2);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
