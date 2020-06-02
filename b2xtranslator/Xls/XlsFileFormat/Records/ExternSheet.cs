
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;
using b2xtranslator.xls.XlsFileFormat.Records;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.ExternSheet)] 
    public class ExternSheet : BiffRecord
    {
        public const RecordType ID = RecordType.ExternSheet;

        public ushort cXTI;

        public ushort[] iSUPBOOK; 

        public ushort[] itabFirst;

        public ushort[] itabLast;

        public List<XTI> rgXTI;

        public ExternSheet(ushort cXTI, List<XTI> rgXTI) : base (RecordType.ExternSheet, 0)
        {
            if (rgXTI.Count != cXTI)
            {
                throw new ArgumentException("cXTI must match the number of XTI entries");
            }

            this.cXTI = cXTI;

            this.iSUPBOOK = new ushort[this.cXTI];
            this.itabFirst = new ushort[this.cXTI];
            this.itabLast = new ushort[this.cXTI];
            this.rgXTI = rgXTI;

            int offset = 0;
            foreach (var xti in rgXTI)
            {
                this.iSUPBOOK[offset] = xti.iSupBook;
                this.itabFirst[offset] = xti.itabFirst;
                this.itabLast[offset] = xti.itabLast;
                offset += 1;
            }

            _length = (uint) (2 + (6 * cXTI));
        }

        public ExternSheet(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            
            this.cXTI = this.Reader.ReadUInt16();

            this.iSUPBOOK = new ushort[this.cXTI];
            this.itabFirst = new ushort[this.cXTI]; 
            this.itabLast = new ushort[this.cXTI];

            rgXTI = new List<XTI>();

            int offset = 0;
            //Sometimes cXTI is invalid - try to continue when that happens
            for (int i = 2; i < Length; i+=6)
            {                
                this.iSUPBOOK[offset] = this.Reader.ReadUInt16();
                this.itabFirst[offset] = this.Reader.ReadUInt16();
                this.itabLast[offset] = this.Reader.ReadUInt16(); 

                rgXTI.Add(new XTI(this.iSUPBOOK[offset], this.itabFirst[offset], this.itabLast[offset]));
                offset += 1;
            }

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }

        public override byte[] GetBytes()
        {
            using (BinaryWriter bw = new BinaryWriter(new MemoryStream()))
            {
                bw.Write(GetHeaderBytes());

                bw.Write(Convert.ToUInt16(this.cXTI));

                foreach (var xti in rgXTI)
                {
                    bw.Write(Convert.ToUInt16(xti.iSupBook));
                    bw.Write(Convert.ToUInt16(xti.itabFirst));
                    bw.Write(Convert.ToUInt16(xti.itabLast));
                }

                return bw.GetBytesWritten();
            }
        }
    }
}
