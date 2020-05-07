
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This class is used to read data from a NUMBER BiffRecord 
    /// </summary>
    [BiffRecord(RecordType.Number)] 
    public class Number : AbstractCellContent
    {
        public const RecordType ID = RecordType.Number;

        /// <summary>
        /// The floating point number
        /// </summary>
        public double? num;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Record ID - Recordtype</param>
        /// <param name="length">The recordlegth</param>
        public Number(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // NOTE: cell fields are parsed by base class

            this.num = new ChartNumNillable(reader).value;
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
