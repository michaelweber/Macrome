using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// PRECISION: Precision (0Eh)
    /// 
    /// The PRECISION record stores the Precision As Displayed option from the Options dialog box, Calculation tab.
    /// </summary>
    [BiffRecord(RecordType.CalcPrecision)] 
    public class CalcPrecision : BiffRecord
    {
        public const RecordType ID = RecordType.CalcPrecision;

        /// <summary>
        /// =0 if Precision As Displayed  option is selected
        /// </summary>
        public ushort fFullPrec;

        public CalcPrecision(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fFullPrec = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
