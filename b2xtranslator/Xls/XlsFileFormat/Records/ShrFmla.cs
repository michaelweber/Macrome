
using System;
using System.Collections.Generic;
using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.ShrFmla)] 
    public class ShrFmla : BiffRecord
    {
        public const RecordType ID = RecordType.ShrFmla;

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
        public ushort colFirst;

        /// <summary>
        /// Colnumber 
        /// </summary>
        public ushort colLast;


        /// <summary>
        /// length of the formular data !!
        /// </summary>
        public ushort cce;

        /// <summary>
        /// LinkedList with the Ptg records !!
        /// </summary>
        public Stack<AbstractPtg> ptgStack; 

        public ShrFmla(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.rwFirst = reader.ReadUInt16();
            this.rwLast = reader.ReadUInt16();
            this.colFirst = (ushort)reader.ReadByte();
            this.colLast = (ushort)reader.ReadByte();

            // read two unnessesary bytes 
            reader.ReadUInt16();

            this.cce = reader.ReadUInt16();
            this.ptgStack = new Stack<AbstractPtg>();
            // reader.ReadBytes(this.cce);

            long oldStreamPosition = this.Reader.BaseStream.Position;
            try
            {
                this.ptgStack = ExcelHelperClass.getFormulaStack(this.Reader, this.cce);
            }
            catch (Exception ex)
            {
                this.Reader.BaseStream.Seek(oldStreamPosition, System.IO.SeekOrigin.Begin);
                this.Reader.BaseStream.Seek(this.cce, System.IO.SeekOrigin.Current);
                TraceLogger.Error(ex.StackTrace); 

            }

            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
