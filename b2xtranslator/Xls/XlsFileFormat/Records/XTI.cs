using System;
using System.Collections.Generic;
using System.Text;

namespace b2xtranslator.xls.XlsFileFormat.Records
{
    public class XTI
    {
        public ushort iSupBook;
        public ushort itabFirst;
        public ushort itabLast;

        public XTI(int iSupBook, int itabFirst, int itabLast)
        {
            this.iSupBook =  (ushort)iSupBook;
            this.itabFirst = (ushort)itabFirst;
            this.itabLast =  (ushort)itabLast;
        }
    }
}
