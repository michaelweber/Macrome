using System;

namespace b2xtranslator.Tools
{
    public class RGBColor
    {
        public enum ByteOrder 
        {
            RedFirst,
            RedLast
        }

        public byte Red;
        public byte Green;
        public byte Blue;
        public byte Alpha;
        public string SixDigitHexCode;
        public string EightDigitHexCode;

        public RGBColor(int cv, ByteOrder order)
        {
            var bytes = System.BitConverter.GetBytes(cv);

            if(order == ByteOrder.RedFirst)
            {
                //R
                this.Red = bytes[0];
                this.SixDigitHexCode = string.Format("{0:x2}", this.Red);
                //G
                this.Green = bytes[1];
                this.SixDigitHexCode += string.Format("{0:x2}", this.Green);
                //B
                this.Blue = bytes[2];
                this.SixDigitHexCode += string.Format("{0:x2}", this.Blue);
                this.EightDigitHexCode = this.SixDigitHexCode;
                //Alpha
                this.Alpha = bytes[3];
                this.EightDigitHexCode += string.Format("{0:x2}", this.Alpha);
            }
            else if (order == ByteOrder.RedLast)
            {
                //R
                this.Red = bytes[2];
                this.SixDigitHexCode = string.Format("{0:x2}", this.Red);
                //G
                this.Green = bytes[1];
                this.SixDigitHexCode += string.Format("{0:x2}", this.Green);
                //B
                this.Blue = bytes[0];
                this.SixDigitHexCode += string.Format("{0:x2}", this.Blue);
                this.EightDigitHexCode = this.SixDigitHexCode;
                //Alpha
                this.Alpha = bytes[3];
                this.EightDigitHexCode += string.Format("{0:x2}", this.Alpha);
            }

        }
    }
}
