
using System.Collections.Generic;
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    [BiffRecord(RecordType.Palette)] 
    public class Palette : BiffRecord
    {
        public const RecordType ID = RecordType.Palette;

        public int ccv;

        public List<RGBColor> rgbColorList; 

        public Palette(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            this.rgbColorList = new List<RGBColor>(); 
            this.ccv = reader.ReadUInt16();

            for (int i = 0; i < this.ccv; i++)
            {
                var color = new RGBColor(reader.ReadInt32(), RGBColor.ByteOrder.RedFirst);
                this.rgbColorList.Add(color); 
            }

        }
    }
}
