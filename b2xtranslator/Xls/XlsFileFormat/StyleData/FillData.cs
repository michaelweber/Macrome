

namespace b2xtranslator.Spreadsheet.XlsFileFormat.StyleData
{
    public class FillData : System.Object
    {
        /// <summary>
        /// Type from this filldata object 
        /// </summary>
        private StyleEnum fillPatern;
        public StyleEnum Fillpatern
        {
            get { return this.fillPatern; }
        }

        /// <summary>
        /// Foreground Color 
        /// </summary>
        private int icvFore;
        public int IcvFore
        {
            get { return this.icvFore; }
        }

        /// <summary>
        /// Background color 
        /// </summary>
        private int icvBack;
        public int IcvBack
        {
            get { return this.icvBack; }
        }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="fillpat">Fill Patern</param>
        /// <param name="icvFore">Foreground Color</param>
        /// <param name="icvBack">Background Color</param>
        public FillData(StyleEnum fillpat, int icvFore, int icvBack)
        {
            this.fillPatern = fillpat;
            this.icvFore = icvFore;
            this.icvBack = icvBack;
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
            var fd = obj as FillData;
            if ((System.Object)fd == null)
            {
                return false;
            }

            // Return true if the fields match:
            return (this.fillPatern == fd.fillPatern) && (this.icvBack == fd.icvBack) && (this.icvFore == fd.icvFore);
        }

        /// <summary>
        /// Equals Method
        /// </summary>
        /// <param name="fd"></param>
        /// <returns></returns>
        public bool Equals(FillData fd)
        {
            // If parameter is null return false:
            if ((object)fd == null)
            {
                return false;
            }

            // Return true if the fields match:
            return (this.fillPatern == fd.fillPatern) && (this.icvBack == fd.icvBack) && (this.icvFore == fd.icvFore);
        }

        /// <summary>
        /// Simple toString method
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return "Fillvalue: " + this.fillPatern.ToString() + "   FG: " + this.icvFore.ToString() + "  BG: " + this.icvBack.ToString();
        }

    }
}
