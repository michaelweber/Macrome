using System.Collections.Generic;
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.Array)] 
    public class ARRAY : BiffRecord
    {
        public const RecordType ID = RecordType.Array;
        /// <summary>
        /// Rownumber 
        /// </summary>
        public ushort rwFirst;

        /// <summary>
        /// Rownumber 
        /// </summary>
        public ushort rwLast;

        /// <summary>
        /// Colnumber 
        /// </summary>
        public byte colFirst;

        /// <summary>
        /// Colnumber 
        /// </summary>
        public byte colLast;

        /// <summary>
        /// option flags 
        /// </summary>
        public ushort grbit;

        /// <summary>
        /// used for performance reasons only 
        /// can be ignored 
        /// </summary>
        public uint chn;

        /// <summary>
        /// length of the formular data !!
        /// </summary>
        public ushort cce;

        /// <summary>
        /// LinkedList with the Ptg records !!
        /// </summary>
        public Stack<AbstractPtg> ptgStack;

        public ARRAY(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.rwFirst = reader.ReadUInt16();
            this.rwLast = reader.ReadUInt16();
            this.colFirst = reader.ReadByte();
            this.colLast = reader.ReadByte();
            this.grbit = reader.ReadUInt16();
            this.chn = reader.ReadUInt32(); // this is used for performance reasons only 
            this.cce = reader.ReadUInt16();
            this.ptgStack = new Stack<AbstractPtg>();
            // reader.ReadBytes(this.cce);

            this.ptgStack = ExcelHelperClass.getFormulaStack(this.Reader, this.cce); 
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
