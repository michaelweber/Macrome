using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.SBaseRef)]
    public class SBaseRef : BiffRecord
    {
        /// <summary>
        ///  A RwU that specifies the zero-based index of the first row in the range. <br/>
        ///  MUST be less than or equal to rwLast.
        /// </summary>
        public ushort rwFirst;

        /// <summary>
        /// A RwU that specifies the zero-based index of the last row in the range. <br/>
        /// MUST be greater than or equal to rwFirst.
        /// </summary>
        public ushort rwLast;

        /// <summary>
        /// A ColU that specifies the zero-based index of the first column in the range.<br/> 
        /// MUST be less than or equal to colLast.<br/>
        /// MUST be less than or equal to 0x00FF.
        /// </summary>
        public ushort colFirst;

        /// <summary>
        /// A ColU that specifies the zero-based index of the last column in the range. <br/>
        /// MUST be greater than or equal to colFirst.<br/>
        /// MUST be less than or equal to 0x00FF.
        /// </summary>
        public ushort colLast;

        public SBaseRef(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            this.rwFirst = reader.ReadUInt16();
            this.rwLast = reader.ReadUInt16();
            this.colFirst = reader.ReadUInt16();
            this.colLast = reader.ReadUInt16();
        }
    }
}
