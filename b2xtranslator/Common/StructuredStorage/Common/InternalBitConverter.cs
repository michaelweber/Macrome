using System;
using System.Collections.Generic;
using System.Text;

namespace b2xtranslator.StructuredStorage.Common
{
    /// <summary>
    /// Wrapper of the class BitConverter in order to support big endian
    /// Author: math
    /// </summary>
    internal class InternalBitConverter
    {

        private bool _IsLittleEndian = true;


        internal InternalBitConverter(bool isLittleEndian)
        {
            this._IsLittleEndian = isLittleEndian;
        }


        internal ulong ToUInt64(byte[] value)
        {
            if (BitConverter.IsLittleEndian ^ this._IsLittleEndian)
            {
                Array.Reverse(value);
            }
            return BitConverter.ToUInt64(value, 0);
        }


        internal uint ToUInt32(byte[] value)
        {
            if (BitConverter.IsLittleEndian ^ this._IsLittleEndian)
            {
                Array.Reverse(value);
            }
            return BitConverter.ToUInt32(value, 0);
        }


        internal ushort ToUInt16(byte[] value)
        {
            if (BitConverter.IsLittleEndian ^ this._IsLittleEndian)
            {
                Array.Reverse(value);
            }
            return BitConverter.ToUInt16(value, 0);
        }


        internal string ToString(byte[] value)
        {
            if (BitConverter.IsLittleEndian ^ this._IsLittleEndian)
            {
                Array.Reverse(value);
            }

            var enc = new UnicodeEncoding();            
            string result = enc.GetString(value);
            if (result.Contains("\0"))
            {
                result = result.Remove(result.IndexOf("\0"));
            }
            return result;
        }


        internal byte[] getBytes(ushort value)
        {
            var result = BitConverter.GetBytes(value);

            if (BitConverter.IsLittleEndian ^ this._IsLittleEndian)
            {
                Array.Reverse(result);
            }
            return result;
        }


        internal byte[] getBytes(uint value)
        {
            var result = BitConverter.GetBytes(value);

            if (BitConverter.IsLittleEndian ^ this._IsLittleEndian)
            {
                Array.Reverse(result);
            }
            return result;
        }


        internal byte[] getBytes(ulong value)
        {
            var result = BitConverter.GetBytes(value);

            if (BitConverter.IsLittleEndian ^ this._IsLittleEndian)
            {
                Array.Reverse(result);
            }
            return result;
        }

        internal List<byte> getBytes(List <uint> input)
        {
            var output = new List<byte>();

            foreach (uint entry in input)
	        {
                output.AddRange(getBytes(entry));
            }
            return output;
        }
    }
}
