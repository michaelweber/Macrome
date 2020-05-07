
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// DSF: Double Stream File (161h)
    /// 
    /// The DSF  record stores a flag that indicates if the workbook is a double stream file.
    /// </summary>
    [BiffRecord(RecordType.DSF)] 
    public class DSF : BiffRecord
    {
        public const RecordType ID = RecordType.DSF;

        /// <summary>
        /// =1 if the workbook is a double stream file
        /// </summary>
        public ushort fDSF;

        public DSF(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.fDSF = reader.ReadUInt16();
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
