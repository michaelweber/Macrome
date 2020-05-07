
using System.Collections.Generic;
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.MulRk)] 
    public class MulRk : BiffRecord
    {
        public const RecordType ID = RecordType.MulRk;

        /// <summary>
        /// Row 
        /// </summary>
        public ushort rw;

        /// <summary>
        /// First column number 
        /// </summary>
        public ushort colFirst;        

        /// <summary>
        /// The last affected column 
        /// </summary>
        public ushort colLast;         

        /// <summary>
        /// List with format indexes 
        /// </summary>
        public List<ushort> ixfe;      // List records 

        /// <summary>
        /// List with the numbers 
        /// </summary>
        public List<double> rknumber;  // List double values 


        /// <summary>
        /// Ctor 
        /// </summary>
        /// <param name="reader">Streamreader</param>
        /// <param name="id">Record ID - Recordtype</param>
        /// <param name="length">The recordlegth</param>
        public MulRk(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            this.ixfe = new List<ushort>();
            this.rknumber = new List<double>(); 

            // count records - 6 standard non variable values !!! 
            int count = (int)(this.Length - 6) / 6 ;
            this.rw = reader.ReadUInt16();
            this.colFirst = reader.ReadUInt16();
            for (int i = 0; i < count; i++)
            {
                this.ixfe.Add(reader.ReadUInt16());
                var buffer = reader.ReadBytes(4);

                this.rknumber.Add(ExcelHelperClass.NumFromRK(buffer)); 
            }
            this.colLast = reader.ReadUInt16(); 
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
       }
    }
}
