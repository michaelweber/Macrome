

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// Specifies an approximation of a real number, where the approximation has a fixed number of digits after the radix point. 
    /// 
    /// This type is specified in [MS-OSHARED] section 2.2.1.6.
    /// 
    /// Value of the real number = Integral + ( Fractional / 65536.0 ) 
    /// 
    /// Integral (2 bytes): A signed integer that specifies the integral part of the real number. 
    /// Fractional (2 bytes): An unsigned integer that specifies the fractional part of the real number.
    /// </summary>
    public class FixedPointNumber
    {
        private short integral;
        private ushort fractional;

        public FixedPointNumber(short integral, ushort fractional)
        {
            this.integral = integral;
            this.fractional = fractional;
        }

        public FixedPointNumber(IStreamReader reader)
        {
            // DEVIATION: The order of fractional and integral part is different as specified.
            this.fractional = reader.ReadUInt16();
            this.integral = reader.ReadInt16();
        }

        public double Value
        {
            get
            {
                return (double)this.integral + (double)this.fractional / 65536.0d;
            }
        }
    }
}
