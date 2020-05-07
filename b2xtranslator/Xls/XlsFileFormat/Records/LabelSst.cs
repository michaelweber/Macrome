
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This class extracts the data from a LABELSST Record
    /// This record describes a cell that contains a string constant from the shared string table 
    /// </summary>
    [BiffRecord(RecordType.LabelSst)] 
    public class LabelSst : BiffRecord
    {
        public const RecordType ID = RecordType.LabelSst;

        /// <summary>
        /// Some public attributes to store the data from this record 
        /// </summary>

        /// <summary>
        /// Rownumber 
        /// </summary>
        public ushort rw;      

        /// <summary>
        /// Colnumber 
        /// </summary>
        public ushort col;     
        
        /// <summary>
        /// index to the XF record 
        /// </summary>
        public ushort ixfe;    

        /// <summary>
        /// index into the SST record  
        /// </summary>
        public uint isst;     

        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="id"></param>
        /// <param name="length"></param>
        public LabelSst(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            this.rw = reader.ReadUInt16();
            this.col = reader.ReadUInt16();
            this.ixfe = reader.ReadUInt16();
            this.isst = reader.ReadUInt32();
           
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
