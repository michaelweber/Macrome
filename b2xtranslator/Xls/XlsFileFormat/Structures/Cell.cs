using System;
using System.Collections.Generic;
using System.Text;

namespace b2xtranslator.xls.XlsFileFormat.Structures
{
    public class Cell
    {
        public UInt16 Rw;
        public UInt16 Col;
        public UInt16 IXFe;

        /// <summary>
        /// Defines a cell structure used by many Excel BIFF Records
        ///
        /// rw and col are straightforward, but ixfe is an offset into the
        /// XFIndex, described at https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/0b683865-eeee-4621-a3d8-b18d20a2afd9
        /// 15 is the default cell style
        /// </summary>
        /// <param name="rw"></param>
        /// <param name="col"></param>
        /// <param name="ixfe"></param>
        public Cell(int rw, int col, int ixfe = 15)
        {

            Rw = Convert.ToUInt16(rw);
            Col = Convert.ToUInt16(col);
            IXFe = Convert.ToUInt16(ixfe);
        }
    }
}
