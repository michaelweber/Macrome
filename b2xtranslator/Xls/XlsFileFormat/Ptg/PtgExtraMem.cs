using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.xls.XlsFileFormat.Ptg
{
    public class PtgExtraMem
    {
        public uint Length;

        public IStreamReader reader;
        public ushort count;

        public List<Ref8U> array;



        public PtgExtraMem(IStreamReader reader)
        {
            this.reader = reader;

            this.count = this.reader.ReadUInt16();

            this.Length = (uint) (2 + 8 * count);

            array = new List<Ref8U>(this.count);

            for (int i = 0; i < count; i += 1)
            {
                array.Add(new Ref8U(reader));
            }
        }
    }
}
