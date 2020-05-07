using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using b2xtranslator.Tools;

namespace b2xtranslator.xls.XlsFileFormat.Structures
{
    public class RgceLoc
    {
        /// row (2 bytes):  An RwU that specifies the row coordinate of the cell reference.
        ///     rw (2 bytes): An unsigned integer that specifies the zero-based index of a row in the sheet that contains this structure.
        /// column(2 bytes) :  A ColRelU that specifies the column coordinate of the cell reference and relative reference information.
        ///     col (14 bits): An unsigned integer that specifies the zero-based index of a column in the sheet that contains this structure.  MUST be less than or equal to 0x00FF.
        ///     A - colRelative(1 bit) : A bit that specifies whether col is a relative reference.
        ///     B - rowRelative(1 bit) : A bit that specifies whether a row index corresponding to col in the structure containing this structure is a relative reference.

        private ushort _rw;
        private ushort _col;
        private bool _colRelative;
        private bool _rwRelative;


        public RgceLoc(ushort rw, ushort col, bool colRelative, bool rwRelative)
        {
            _rw = rw;
            _col = col;
            _colRelative = colRelative;
            _rwRelative = rwRelative;
        }

        public byte[] Bytes
        {
            get
            {
                MemoryStream ms = new MemoryStream();
                BinaryWriter bw = new BinaryWriter(ms);
                bw.Write(Convert.ToUInt16(_rw));
                ushort colWithFlags = Convert.ToUInt16(_col);
                if (_colRelative) colWithFlags = (ushort) (colWithFlags | 0x4000);
                if (_rwRelative) colWithFlags = (ushort) (colWithFlags | 0x8000);
                bw.Write(colWithFlags);
                return bw.GetBytesWritten();
            }
        }
    }
}
