
using System.Collections.Generic;
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.DataContainer;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This class extracts the data from a mergecell biffrecord 
    /// 
    /// MERGECELLS: Merged Cells (E5h) 
    /// This record stores all merged cells. Record Data 
    /// Offset 	Field Name 	Size 	Contents 
    /// 4 	cmcs 	2 	Count of REF structures
    /// 
    /// The REF structure has the following fields. 
    /// Offset 	Field Name 	Size 	Contents 
    /// 8 	rwFirst 	2 	The first row of the range associated with the record 
    /// 10 	rwLast 	2 	The last row of the range associated with the record 
    /// 12 	colFirst 	2 	The first column of the range associated with the record 
    /// 14 	colLast 	2 	The last column of the range associated with the record 
    /// </summary>
    [BiffRecord(RecordType.MergeCells)] 
    public class MergeCells : BiffRecord
    {
        public const RecordType ID = RecordType.MergeCells;

        /// <summary>
        /// List with datarecords 
        /// </summary>
        public List<MergeCellData> mergeCellDataList;

        /// <summary>
        /// Count REF structures 
        /// </summary>
        public ushort cmcs;

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="id"></param>
        /// <param name="length"></param>
        public MergeCells(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            this.mergeCellDataList = new List<MergeCellData>(); 
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.cmcs = this.Reader.ReadUInt16();

            for (int i = 0; i < this.cmcs; i++)
            {
                var mcd = new MergeCellData();
                mcd.rwFirst = this.Reader.ReadUInt16();
                mcd.rwLast = this.Reader.ReadUInt16();
                mcd.colFirst = this.Reader.ReadUInt16();
                mcd.colLast = this.Reader.ReadUInt16();
                this.mergeCellDataList.Add(mcd); 
            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
