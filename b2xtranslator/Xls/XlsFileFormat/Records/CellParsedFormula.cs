using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Tools;

namespace b2xtranslator.xls.XlsFileFormat.Records
{
    public class CellParsedFormula
    {
        public UInt16 cce;
        public byte[] rgce;
        public byte[] rgcb;

        public Stack<AbstractPtg> PtgStack;

        public CellParsedFormula(Stack<AbstractPtg> ptgStack)
        {
            rgce = PtgHelper.GetBytes(ptgStack);
            PtgStack = ptgStack;
            rgcb = new byte[] { };
            cce = Convert.ToUInt16(rgce.Length);
        }

        public byte[] Bytes
        {
            get
            {
                MemoryStream ms = new MemoryStream();
                BinaryWriter bw = new BinaryWriter(ms);
                bw.Write(Convert.ToUInt16(cce));
                bw.Write(rgce);
                bw.Write(rgcb);
                return bw.GetBytesWritten();
            }
        }

    }
}
