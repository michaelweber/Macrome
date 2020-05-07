namespace b2xtranslator.Tools
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
        public ushort Integral;
        public ushort Fractional;

        public FixedPointNumber(ushort integral, ushort fractional)
        {
            this.Integral = integral;
            this.Fractional = fractional;
        }

        public FixedPointNumber(uint value)
        {
            var bytes = System.BitConverter.GetBytes(value);
            this.Integral = System.BitConverter.ToUInt16(bytes, 0);
            this.Fractional = System.BitConverter.ToUInt16(bytes, 2);
        }

        public FixedPointNumber(byte[] bytes)
        {
            this.Integral = System.BitConverter.ToUInt16(bytes, 0);
            this.Fractional = System.BitConverter.ToUInt16(bytes, 2);
        }

        //public FixedPointNumber(IStreamReader reader)
        //{
        //    this.integral = reader.ReadUInt16();
        //    this.fractional = reader.ReadUInt16();
        //}

        public double ToAngle()
        {
            if (this.Fractional != 0)
            {
                // negative angle
                return (this.Fractional - 65536.0);
            }
            else if (this.Integral != 0)
            {
                //positive angle
                return (65536.0 - this.Integral);
            }
            else
            {
                return 0.0;
            }
        }

        public double Value
        {
            get
            {
                return (double)this.Integral + (double)this.Fractional / 65536.0d;
            }
        }
    }
}
