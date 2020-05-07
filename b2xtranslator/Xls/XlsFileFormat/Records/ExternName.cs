
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.ExternName)] 
    public class ExternName : BiffRecord
    {
        public const RecordType ID = RecordType.ExternName;

        public ushort ixals;
        public bool fOle;
        public bool fOleLink;
        public ushort grbit; 
        public string extName;
        public byte cch;

        public ushort cce;

        public string nameDefinition; 

        public ExternName(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.grbit = this.Reader.ReadUInt16();



            this.fOle = Utils.BitmaskToBool(this.grbit, 0x0008);
            this.fOleLink = Utils.BitmaskToBool(this.grbit, 0x0010);

            // this is an external link 
            if (!this.fOle && !this.fOleLink)
            {

                this.ixals = this.Reader.ReadUInt16();
                // read unused 16 bit
                this.Reader.ReadBytes(2);
                this.cch = this.Reader.ReadByte();
                byte firstbyte = this.Reader.ReadByte();
                int firstbit = firstbyte & 0x1;
                for (int i = 0; i < this.cch; i++)
                {
                    if (firstbit == 0)
                    {
                        this.extName += (char)this.Reader.ReadByte();
                        // read 1 byte per char 
                    }
                    else
                    {
                        // read two byte per char 
                        this.extName += System.BitConverter.ToChar(this.Reader.ReadBytes(2), 0);
                    }
                }
                this.cce = this.Reader.ReadUInt16();
                this.Reader.ReadBytes(this.cce); 

            }
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
