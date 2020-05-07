
using System;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.IO;
using System.Text;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.String)] 
    public class STRING : BiffRecord
    {
        public const RecordType ID = RecordType.String;

        public string value;

        public int cch;

        public int grbit;

        private bool isUnicode;

        public STRING(string contents, bool isUnicode) : base(RecordType.String, 0)
        {
            value = contents;
            grbit = (int) Utils.BoolToBitmask(isUnicode, 0x01);
            cch = contents.Length;
            this.isUnicode = isUnicode;
            if (isUnicode)
            {
                _length = (uint) (3 + cch * 2);
            }
            else
            {
                _length = (uint) (3 + cch);
            }
        }

        public STRING(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.cch = reader.ReadUInt16();

            this.grbit = reader.ReadByte();

            this.value = ExcelHelperClass.getStringFromBiffRecord(reader, this.cch, this.grbit); 
	

            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }

        public override byte[] GetBytes()
        {
            using (BinaryWriter bw = new BinaryWriter(new MemoryStream()))
            {
                bw.Write(GetHeaderBytes());
                bw.Write(Convert.ToUInt16(cch));
                bw.Write(Convert.ToByte(grbit));

                if (isUnicode)
                {
                    bw.Write(Encoding.Unicode.GetBytes(this.value));
                }
                else
                {
                    bw.Write(Encoding.GetEncoding(1252).GetBytes(this.value));
                }

                return bw.GetBytesWritten();
            }
        }
    }
}
