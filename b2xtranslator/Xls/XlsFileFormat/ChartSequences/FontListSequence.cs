using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    public class FontFbiGroup
    {
        private Font _font;
        private Fbi _fbi;

        public FontFbiGroup(Font font, Fbi fbi)
        {
            this._font = font;
            this._fbi = fbi;
        }

        public Font Font
        {
            get { return this._font; }
            set { this._font = value; }
        }

        public Fbi Fbi
        {
            get { return this._fbi; }
            set { this._fbi = value; }
        }
    }

    public class FontListSequence : BiffRecordSequence
    {
        public FrtFontList FrtFontList;

        //public StartObject StartObject;

        public List<FontFbiGroup> Fonts;

        //public EndObject EndObject;

        public FontListSequence(IStreamReader reader) : base(reader)
        {
            //FrtFontList 
            this.FrtFontList = (FrtFontList)BiffRecord.ReadRecord(reader);

            //StartObject 
            //this.StartObject = (StartObject)BiffRecord.ReadRecord(reader);
            
            //*(Font [Fbi]) 
            this.Fonts = new List<FontFbiGroup>();
            while (BiffRecord.GetNextRecordType(reader) == RecordType.Font)
            {
                var font = (Font)BiffRecord.ReadRecord(reader);
                Fbi fbi = null;
                if (BiffRecord.GetNextRecordType(reader) == RecordType.Fbi)
                {
                    fbi = (Fbi)BiffRecord.ReadRecord(reader);
                }
                this.Fonts.Add(new FontFbiGroup(font, fbi));
            }

            //EndObject
            //this.EndObject = (EndObject)BiffRecord.ReadRecord(reader);
        }
    }
}
