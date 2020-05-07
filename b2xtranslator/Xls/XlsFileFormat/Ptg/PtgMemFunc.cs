using System;
using System.Collections.Generic;
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgMemFunc : AbstractPtg
    {
        public const PtgNumber ID = PtgNumber.PtgMemFunc;

        /// <summary>
        /// LinkedList with the Ptg records !!
        /// </summary>
        public Stack<AbstractPtg> ptgStack;

        public PtgMemFunc(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);

            this.Data = "";

            this.type = PtgType.Operand;
            this.popSize = 1;

            int cce = reader.ReadUInt16();
            this.Length = (uint)(3 + cce);

            long oldStreamPosition = this.Reader.BaseStream.Position;



            try
            {
                this.ptgStack = ExcelHelperClass.getFormulaStack(this.Reader, (ushort)cce);
            }
            catch (Exception ex)
            {
                this.Reader.BaseStream.Seek(oldStreamPosition, System.IO.SeekOrigin.Begin);
                this.Reader.BaseStream.Seek(cce, System.IO.SeekOrigin.Current);
                TraceLogger.Error(ex.StackTrace);
            }



        }
    }
}
