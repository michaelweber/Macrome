using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.BoolErr)] 
    public class BoolErr : AbstractCellContent
    {
        public const RecordType ID = RecordType.BoolErr;

        /// <summary>
        /// An unsigned integer that specifies either a Boolean value or an error value, depending on the value of fError.
        /// </summary>
        public byte bBoolErr;

        /// <summary>
        /// A Boolean that specifies whether bBoolErr contains an error code or a Boolean value.
        /// </summary>
        public bool fError;


        public BoolErr(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.bBoolErr = reader.ReadByte();
            this.fError = reader.ReadByte() == 0x1;
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
