using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This class is used to read data from a NUMBER BiffRecord 
    /// </summary>
    public class AbstractCellContent : BiffRecord
    {
        /// <summary>
        /// Row 
        /// </summary>
        public ushort rw;
        /// <summary>
        /// Column
        /// </summary>
        public ushort col;
        /// <summary>
        /// Index to the XF Record 
        /// </summary>
        public ushort ixfe;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Record ID - Recordtype</param>
        /// <param name="length">The recordlegth</param>
        public AbstractCellContent(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            this.rw = reader.ReadUInt16();
            this.col = reader.ReadUInt16();
            this.ixfe = reader.ReadUInt16();
        }
    }
}
