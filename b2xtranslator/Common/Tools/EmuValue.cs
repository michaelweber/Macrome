using System;
using System.Globalization;

namespace b2xtranslator.Tools
{

    public class EmuValue
    {
        public int Value;

        /// <summary>
        /// Creates a new EmuValue for the given value.
        /// </summary>
        /// <param name="value"></param>
        public EmuValue(int value)
        {
            this.Value = value;
        }

        /// <summary>
        /// Converts the EMU to pt
        /// </summary>
        /// <returns></returns>
        public double ToPoints()
        {
            return this.Value / 12700;
        }

        /// <summary>
        /// Converts the EMU to twips
        /// </summary>
        /// <returns></returns>
        public double ToTwips()
        {
            return this.Value / 635;
        }

        public double ToCm()
        {
            return this.Value / 36000.0;
        }

        public double ToMm()
        {
            return ToCm() * 10.0;
        }

        /// <summary>
        /// returns the original value as string 
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Convert.ToString(this.Value, CultureInfo.GetCultureInfo("en-US")); 
        }
    }
}
