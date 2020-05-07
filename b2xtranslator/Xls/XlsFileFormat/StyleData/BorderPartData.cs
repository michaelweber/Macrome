


namespace b2xtranslator.Spreadsheet.XlsFileFormat.StyleData
{
    public class BorderPartData
    {
        public ushort style;
        public int colorId;

        public BorderPartData(ushort style, int colorId)
        {
            this.style = style; 
            this.colorId = colorId; 
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
            var bpd = obj as BorderPartData;
            if ((System.Object)bpd == null)
            {
                return false;
            }

            // Return true if the fields match:
            return (this.colorId == bpd.colorId) && (this.style == bpd.style);
        }

        /// <summary>
        /// Equals Method
        /// </summary>
        /// <param name="fd"></param>
        /// <returns></returns>
        public bool Equals(BorderPartData bpd)
        {
            // If parameter is null return false:
            if ((object)bpd == null)
            {
                return false;
            }

            // Return true if the fields match:
            return (this.colorId == bpd.colorId) && (this.style == bpd.style);
        }
    }


}
