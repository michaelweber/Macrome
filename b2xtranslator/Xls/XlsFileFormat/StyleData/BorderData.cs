


namespace b2xtranslator.Spreadsheet.XlsFileFormat.StyleData
{
    public class BorderData
    {
        public BorderPartData top;
        public BorderPartData bottom;
        public BorderPartData left;
        public BorderPartData right;
        public BorderPartData diagonal;

        public ushort diagonalValue; 

        public BorderData()
        {

            this.diagonalValue = 0; 
        }

        /// <summary>
        /// Equals Method 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public override bool Equals(System.Object obj)
        {
            // If parameter is null return false.
            if (obj == null)
            {
                return false;
            }

            // If parameter cannot be cast to FillDataList return false.
            var bd = obj as BorderData;
            if ((System.Object)bd == null)
            {
                return false;
            }

            if (this.top.Equals(bd.top) && this.bottom.Equals(bd.bottom) && this.left.Equals(bd.left)
                && this.right.Equals(bd.right) && this.diagonal.Equals(bd.diagonal) 
                && this.diagonalValue == bd.diagonalValue)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Equals Method
        /// </summary>
        /// <param name="fd"></param>
        /// <returns></returns>
        public bool Equals(BorderData bd)
        {
            // If parameter is null return false:
            if ((object)bd == null)
            {
                return false;
            }

            if (this.top.Equals(bd.top) && this.bottom.Equals(bd.bottom) && this.left.Equals(bd.left)
                && this.right.Equals(bd.right) && this.diagonal.Equals(bd.diagonal)
                && this.diagonalValue == bd.diagonalValue)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

    }
}
