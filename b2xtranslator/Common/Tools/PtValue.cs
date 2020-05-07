using System;
using System.Globalization;

namespace b2xtranslator.Tools
{

    public class PtValue
    {
        public double Value;

        /// <summary>
        /// Creates a new PtValue for the given value.
        /// </summary>
        /// <param name="value"></param>
        public PtValue(double value)
        {
            this.Value = value;
        }

        /// <summary>
        /// Converts the EMU to pt
        /// </summary>
        /// <returns></returns>
        public double ToPoints()
        {
            return this.Value;
        }

        /// <summary>
        /// Converts the pt value to EMU
        /// </summary>
        /// <returns></returns>
        public int ToEmu()
        {
            return (int)((360000 * 2.54 * this.Value) / 72.0);
        }

        /// <summary>
        /// Converts the pt value to cm
        /// </summary>
        /// <returns></returns>
        public double ToCm()
        {
            return (2.54 * this.Value) / 72.0;
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
