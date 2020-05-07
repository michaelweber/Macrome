
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;


namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.HLink)] 
    public class HLink : BiffRecord
    {
        public const RecordType ID = RecordType.HLink;

        public ushort rwFirst;
        public ushort rwLast;
        public ushort colFirst;
        public ushort colLast;
        public uint streamVersion;

        public bool hlstmfIsAbsolute;

        public byte[] hlinkClsid; 

        public string displayName;
        public string targetFrameName;
        public string monikerString;
        public string location;
        public byte[] guid;
        public byte[] fileTime; 
        
        
        public HLink(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);
            
            this.rwFirst = this.Reader.ReadUInt16();
            this.rwLast = this.Reader.ReadUInt16();
            this.colFirst = this.Reader.ReadUInt16();
            this.colLast = this.Reader.ReadUInt16();

            this.hlinkClsid = new byte[16];

            // read 16 bytes hlinkClsid
            this.hlinkClsid = this.Reader.ReadBytes(16);
            this.streamVersion = this.Reader.ReadUInt32();

            uint buffer = this.Reader.ReadUInt32();
            bool hlstmfHasMoniker = Utils.BitmaskToBool(buffer, 0x01);
            this.hlstmfIsAbsolute = Utils.BitmaskToBool(buffer, 0x02);
            bool hlstmfSiteGaveDisplayName = Utils.BitmaskToBool(buffer, 0x04);
            bool hlstmfHasLocationStr = Utils.BitmaskToBool(buffer, 0x08);
            bool hlstmfHasDisplayName = Utils.BitmaskToBool(buffer, 0x10);
            bool hlstmfHasGUID = Utils.BitmaskToBool(buffer, 0x20);
            bool hlstmfHasCreationTime = Utils.BitmaskToBool(buffer, 0x40);
            bool hlstmfHasFrameName = Utils.BitmaskToBool(buffer, 0x80);
            bool hlstmfMonikerSavedAsStr = Utils.BitmaskToBool(buffer, 0x100);
            bool hlstmfAbsFromGetdataRel = Utils.BitmaskToBool(buffer, 0x200);

            if (hlstmfHasDisplayName)
            {
                this.displayName = ExcelHelperClass.getHyperlinkStringFromBiffRecord(this.Reader);
            }
            if (hlstmfHasFrameName)
            {
                
                this.targetFrameName = ExcelHelperClass.getHyperlinkStringFromBiffRecord(this.Reader);
            }
            if (hlstmfHasMoniker)
            {
                if (hlstmfMonikerSavedAsStr)
                {
                    this.monikerString = ExcelHelperClass.getHyperlinkStringFromBiffRecord(this.Reader);
                }
                else
                {
                    // OleMoniker 
                    // read monikerClsid
                    uint Part1MonikerClsid = this.Reader.ReadUInt32();
                    ushort Part2MonikerClsid = this.Reader.ReadUInt16();
                    ushort Part3MonikerClsid = this.Reader.ReadUInt16();

                    byte Part4MonikerClsid = this.Reader.ReadByte();
                    byte Part5MonikerClsid = this.Reader.ReadByte();
                    byte Part6MonikerClsid = this.Reader.ReadByte();
                    byte Part7MonikerClsid = this.Reader.ReadByte();
                    byte Part8MonikerClsid = this.Reader.ReadByte();
                    byte Part9MonikerClsid = this.Reader.ReadByte();
                    byte Part10MonikerClsid = this.Reader.ReadByte();
                    byte Part11MonikerClsid = this.Reader.ReadByte(); 

                    // URL Moniker
                    if (Part1MonikerClsid == 0x79EAC9E0)
                    {
                        uint lenght = reader.ReadUInt32();
                        string value = "";
                        // read until the \0 value 

                        do
                        {
                            value += System.BitConverter.ToChar(reader.ReadBytes(2), 0);
                        } while (value[value.Length - 1] != '\0');

                        if (value.Length * 2 != lenght)
                        {
                            // read guid serial version and uriflags 
                            this.Reader.ReadBytes(24);
                        }
                        value = value.Remove(value.Length - 1);
                        this.monikerString = value; 
                    }
                    else if (Part1MonikerClsid == 0x00000303)
                    {
                        ushort cAnti = this.Reader.ReadUInt16();
                        uint ansiLength = this.Reader.ReadUInt32();
                        string ansiPath = "";
                        for (int i = 0; i < ansiLength; i++)
                        {
                            ansiPath += (char)reader.ReadByte();
                            
                        }


                        ansiPath = ansiPath.Remove(ansiPath.Length - 1);
                        ushort endServer = this.Reader.ReadUInt16();
                        ushort versionNumber = this.Reader.ReadUInt16();
                        this.monikerString = ansiPath; 
                        // read 20 unused bytes 
                        this.Reader.ReadBytes(20);
                        uint cbUnicodePathSize = this.Reader.ReadUInt32();
                        //string unicodePath = ""; 

                        if (cbUnicodePathSize != 0)
                        {
                            uint cbUnicodePathBytes = this.Reader.ReadUInt32();
                            ushort usKeyValue = this.Reader.ReadUInt16();

                            string value = "";

                            for (int i = 0; i < cbUnicodePathBytes/2; i++)
                            {
                                value += System.BitConverter.ToChar(reader.ReadBytes(2), 0);
                            }
                            this.monikerString = value; 
                        }


                    }

                    //byte[] monikerClsid = this.Reader.ReadBytes(16);
                    //string monikerid = "";
                    //for (int i = 0; i < monikerClsid.Length; i++)
                    //{
                    //    monikerid = monikerid + monikerClsid[i].ToString(); 
                    //}
                }
            }
            if (hlstmfHasLocationStr)
            {
                this.location = ExcelHelperClass.getHyperlinkStringFromBiffRecord(this.Reader);
            }
            if (hlstmfHasGUID)
            {
                this.guid = this.Reader.ReadBytes(16);
            }
            if (hlstmfHasCreationTime)
            {
                this.fileTime = this.Reader.ReadBytes(8);
            }
        }
    }
}
