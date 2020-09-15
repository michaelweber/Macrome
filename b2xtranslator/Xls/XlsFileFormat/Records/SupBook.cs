
using System;
using System.Diagnostics;
using System.IO;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.SupBook)] 
    public class SupBook : BiffRecord
    {
        public const RecordType ID = RecordType.SupBook;

        public ushort ctab;

        public ushort cch; 

        public string virtpathstring;

        public string[] rgst;

        public bool isvirtpath;
        public bool isexternalworkbookreferencing;
        public bool isselfreferencing;
        public bool isaddinreferencing;
        public bool isunusedsupportinglink; 

        public SupBook(int ctab, int cch) : base (RecordType.SupBook, 4)
        {
            this.ctab = (ushort) ctab;
            this.cch = (ushort) cch;

            this.isselfreferencing = true;
            this.isaddinreferencing = false;
            this.isvirtpath = false;
            this.isexternalworkbookreferencing = false;
            this.isunusedsupportinglink = false;

            // Check cch 
            if (this.cch == 0x0401)
            {
                this.isselfreferencing = true;
            }
            else if (this.cch == 0x3A01)
            {
                this.isaddinreferencing = true;
                //0x0001 to 0x00ff (inclusive)
            }
            else if (this.cch >= 0x0001 && this.cch <= 0x00ff)
            {
                this.isvirtpath = true;
            }
        }

        public SupBook(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.ctab = this.Reader.ReadUInt16();

            this.cch = this.Reader.ReadUInt16();

            this.isselfreferencing = true;
            this.isaddinreferencing = false;
            this.isvirtpath = false;
            this.isexternalworkbookreferencing = false;
            this.isunusedsupportinglink = false; 

            // Check cch 
            if (this.cch == 0x0401)
            {
                this.isselfreferencing = true; 
            }
            else if (this.cch == 0x3A01)
            {
                this.isaddinreferencing = true; 
                //0x0001 to 0x00ff (inclusive)
            }
            else if (this.cch >= 0x0001 && this.cch <= 0x00ff)
            {
                this.isvirtpath = true; 
            }

            if (this.isvirtpath)
            {
                this.virtpathstring = ""; 
                byte firstbyte = this.Reader.ReadByte();
                int firstbit = firstbyte & 0x1;
                for (int i = 0; i < this.cch; i++)
                {
                    if (firstbit == 0)
                    {
                        this.virtpathstring += (char)this.Reader.ReadByte();
                        // read 1 byte per char 
                    }
                    else
                    {
                        // read two byte per char 
                        this.virtpathstring += System.BitConverter.ToChar(this.Reader.ReadBytes(2), 0);
                    }
                }
                this.virtpathstring = ExcelHelperClass.parseVirtualPath(this.virtpathstring);
            }
            
            if (this.virtpathstring != null)
            {
                if (this.virtpathstring.Equals(0x00))
                {
                    this.isselfreferencing = true;
                }
                else if (this.virtpathstring.Equals(0x20))
                {
                    this.isunusedsupportinglink = true;
                }
                else
                {
                    this.isexternalworkbookreferencing = true; 
                }
            }

            if ((this.isexternalworkbookreferencing) || (this.isunusedsupportinglink))
            {
                this.rgst = new string[this.ctab];
                for (int i = 0; i < this.ctab; i++)
                {

                    ushort cch2 = this.Reader.ReadUInt16(); 
                        byte firstbyte = this.Reader.ReadByte();
                        int firstbit = firstbyte & 0x1;
                        for (int j = 0; j < cch2; j++)
                        {
                            if (firstbit == 0)
                            {
                                this.rgst[i] += (char)this.Reader.ReadByte();
                                // read 1 byte per char 
                            }
                            else
                            {
                                // read two byte per char 
                                this.rgst[i] += System.BitConverter.ToChar(this.Reader.ReadBytes(2), 0);
                            }
                        }        
                }
            }
            if (this.virtpathstring != null && this.virtpathstring.Length > 1)
                this.isselfreferencing = false;

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position); 
        }

        public override byte[] GetBytes()
        {
            using (BinaryWriter bw = new BinaryWriter(new MemoryStream()))
            {
                bw.Write(GetHeaderBytes());
                bw.Write(Convert.ToUInt16(this.ctab));
                bw.Write(Convert.ToUInt16(this.cch));

                if (this.isvirtpath && this.ctab == 1 && this.cch < 0xff)
                {
                    XLUnicodeStringNoCch virtPathXLU = new XLUnicodeStringNoCch(virtpathstring);
                    bw.Write(virtPathXLU.Bytes);
                    foreach (var rgstString in this.rgst)
                    {
                        ushort length = Convert.ToUInt16(rgstString.Length);
                        bw.Write(length);
                        XLUnicodeStringNoCch rgstXLU = new XLUnicodeStringNoCch(rgstString);
                        bw.Write(rgstXLU.Bytes);
                    }
                }
                else if (this.isvirtpath || this.isexternalworkbookreferencing || this.isunusedsupportinglink)
                {
                    throw new NotImplementedException();
                }

                return bw.GetBytesWritten();
            }
        }
    }
}
