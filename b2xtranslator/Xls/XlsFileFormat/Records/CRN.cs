
using System;
using System.Collections.Generic;
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.CRN)] 
    public class CRN : BiffRecord
    {
        public const RecordType ID = RecordType.CRN;

        public byte colLast;
        public byte colFirst;

        public ushort rw;

        public List<object> oper; 


        public CRN(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            this.oper = new List<object>(); 
            long endposition = this.Reader.BaseStream.Position + this.Length; 
            this.colLast = this.Reader.ReadByte();
            this.colFirst = this.Reader.ReadByte();
            this.rw = this.Reader.ReadUInt16();

            


            while (this.Reader.BaseStream.Position < endposition)
            {
                byte grbit = this.Reader.ReadByte();
                if (grbit == 0x01)
                {
                    this.oper.Add(this.Reader.ReadDouble()); 
                }
                else if (grbit == 0x02)
                {
                    string Data = "";
                    ushort cch = this.Reader.ReadUInt16();
                    byte firstbyte = this.Reader.ReadByte();
                    int firstbit = firstbyte & 0x1;
                    for (int i = 0; i < cch; i++)
                    {
                        if (firstbit == 0)
                        {
                            Data += (char)this.Reader.ReadByte();
                            // read 1 byte per char 
                        }
                        else
                        {
                            // read two byte per char 
                            Data += System.BitConverter.ToChar(this.Reader.ReadBytes(2), 0);
                        }
                    }

                    this.oper.Add(Data); 
                }
                else if (grbit == 0x00)
                {
                    this.Reader.ReadBytes(8);
                    this.oper.Add(" "); 
                }
                else if (grbit == 0x04)
                {
                    // bool 
                    ushort boolvalue = this.Reader.ReadUInt16();
                    bool value = false;
                    if (boolvalue == 1)
                        value = true;
                    this.oper.Add(value);
                    this.Reader.ReadBytes(6); 
                }
                else if (grbit == 0x10)
                {
                    // Error
                    this.Reader.ReadBytes(8);
                    this.oper.Add("Err"); 
                }

            }
            
            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }
    }
}
