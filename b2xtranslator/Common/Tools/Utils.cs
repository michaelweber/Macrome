using System;
using System.Text;
using System.IO;
using System.Collections;
using System.ComponentModel;
using System.Xml;

namespace b2xtranslator.Tools
{
    public class Utils
    {
        //Microsoft Office 2007 Beta 2 used namespaces that are not valid anymore.
        //This method should be used for all xml-code inside the binary Powerpoint file.
        //e.g. Themes, Layouts
        public static void replaceOutdatedNamespaces(ref XmlNode e)
        {
            string s = replaceOutdatedNamespaces(e.OuterXml);
            var d2 = new XmlDocument();
            d2.LoadXml(s);
            e = d2.DocumentElement;
        }

        public static string replaceOutdatedNamespaces(string input)
        {
            string output = input.Replace("http://schemas.openxmlformats.org/drawingml/2006/3/main", "http://schemas.openxmlformats.org/drawingml/2006/main");
            output = output.Replace("http://schemas.openxmlformats.org/presentationml/2006/3/main", "http://schemas.openxmlformats.org/presentationml/2006/main");
            return output;
        }

        public static string ReadWString(Stream stream)
        {
            var cch = new byte[1];
            stream.Read(cch, 0, cch.Length);

            var chars = new byte[2 * cch[0]];
            stream.Read(chars, 0, chars.Length);

            return Encoding.Unicode.GetString(chars);
        }

        /// <summary>
        /// Read a length prefixed Unicode string from the given stream.
        /// The string must have the following structure:<br/>
        /// byte 1 - 4:         Character count (cch)<br/>
        /// byte 5 - (cch*2)+4: Unicode characters terminated by \0
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static string ReadLengthPrefixedUnicodeString(Stream stream)
        {
            var cchBytes = new byte[4];
            stream.Read(cchBytes, 0, cchBytes.Length);
            int cch = System.BitConverter.ToInt32(cchBytes, 0);

            //dont read the terminating zero
            var stringBytes = new byte[cch*2];
            stream.Read(stringBytes, 0, stringBytes.Length);

            return Encoding.Unicode.GetString(stringBytes, 0, stringBytes.Length-2);
        }

        /// <summary>
        /// Read a length prefixed ANSI string from the given stream.
        /// The string must have the following structure:<br/>
        /// byte 1-4:       Character count (cch)<br/>
        /// byte 5-cch+4:   ANSI characters terminated by \0
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static string ReadLengthPrefixedAnsiString(Stream stream)
        {
            var cchBytes = new byte[4];
            stream.Read(cchBytes, 0, cchBytes.Length);
            int cch = System.BitConverter.ToInt32(cchBytes, 0);

            //dont read the terminating zero
            var stringBytes = new byte[cch];
            stream.Read(stringBytes, 0, stringBytes.Length);

            if (cch > 0)
                return Encoding.ASCII.GetString(stringBytes, 0, stringBytes.Length - 1);
            else
                return null;
        }

        public static string ReadXstz(byte[] bytes, int pos)
        {
            var xstz = new byte[System.BitConverter.ToUInt16(bytes, pos) * 2];
            Array.Copy(bytes, pos + 2, xstz, 0, xstz.Length);
            return Encoding.Unicode.GetString(xstz);
        }

        public static string ReadXst(Stream stream)
        {
            // read the char count
            var cch = new byte[2];
            stream.Read(cch, 0, cch.Length);
            ushort charCount = System.BitConverter.ToUInt16(cch, 0);

            // read the string
            var xst = new byte[charCount * 2];
            stream.Read(xst, 0, xst.Length);
            return Encoding.Unicode.GetString(xst);
        }

        public static string ReadXstz(Stream stream)
        {
            string xst = ReadXst(stream);
            
            //skip the termination
            var termiantion = new byte[2];
            stream.Read(termiantion, 0, termiantion.Length);

            return xst;
        }

        public static string ReadShortXlUnicodeString(Stream stream)
        {
            var cch = new byte[1];
            stream.Read(cch, 0, cch.Length);

            var fHighByte = new byte[1];
            stream.Read(fHighByte, 0, fHighByte.Length);

            int rgbLength = cch[0];
            if (fHighByte[0] >= 0)
            {
                //double byte characters
                rgbLength *= 2;
            }

            var rgb = new byte[rgbLength];
            stream.Read(rgb, 0, rgb.Length);

            if (fHighByte[0] >= 0)
            {
                return Encoding.Unicode.GetString(rgb);
            }
            else
            {
                var enc = Encoding.GetEncoding(1252);
                return enc.GetString(rgb);
            }
        }

        public static int ArraySum(byte[] values)
        {
            int ret = 0;
            foreach (byte b in values)
            {
                ret += b;
            }
            return ret;
        }

        public static bool BitmaskToBool(int value, int mask)
        {
            return ((value & mask) == mask);
        }

        public static bool BitmaskToBool(uint value, uint mask)
        {
            return ((value & mask) == mask);
        }

        public static UInt32 BoolToBitmask(bool value, uint mask)
        {
            return (value) ? mask : 0;
        }

        public static UInt32 ByteToBitmask(byte value, uint mask)
        {
            UInt32 val = Convert.ToUInt32(value);

            while ((mask & 0x1) != 0x1)
            {
                val = val << 1;
                mask = mask >> 1;
            }
            return val;
        }


        public static byte BitmaskToByte(uint value, uint mask)
        {
            value = value & mask;
            while ((mask & 0x1) != 0x1)
            {
                value = value >> 1;
                mask = mask >> 1;
            }
            return Convert.ToByte(value);
        }

        public static int BitmaskToInt(int value, int mask)
        {
            int ret = value & mask;
            //shift for all trailing zeros
            var bits = new BitArray(new int[] { mask });
            foreach (bool bit in bits)
            {
                if (!bit)
                    ret = ret >> 1;
                else
                    break;
            }
            return ret;
        }

        public static int BitmaskToInt32(int value, int mask)
        {
            value = value & mask;
            while ((mask & 0x1) != 0x1)
            {
                value = value >> 1;
                mask = mask >> 1;
            }
            return value;
        }

        public static uint BitmaskToUInt32(uint value, uint mask)
        {
            value = value & mask;
            while ((mask & 0x1) != 0x1)
            {
                value = value >> 1;
                mask = mask >> 1;
            }
            return value;
        }

        public static ushort BitmaskToUInt16(uint value, uint mask)
        {
            return Convert.ToUInt16(BitmaskToUInt32(value, mask));
        }

        public static bool IntToBool(int value)
        {
            if (value == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        public static bool ByteToBool(byte value)
        {
            if (value == 1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static char[] ClearCharArray(char[] values)
        {
            var ret = new char[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                ret[i] = Convert.ToChar(0);
            }
            return ret;
        }

        public static int[] ClearIntArray(int[] values)
        {
            var ret = new int[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                ret[i] = 0;
            }
            return ret;
        }

        public static short[] ClearShortArray(ushort[] values)
        {
            var ret = new short[values.Length];
            for (int i = 0; i < values.Length; i++)
            {
                ret[i] = 0;
            }
            return ret;
        }

        public static uint BitArrayToUInt32(BitArray bits)
        {
            double ret = 0;
            for (int i = 0; i < bits.Count; i++)
            {
                if (bits[i])
                {
                    ret += Math.Pow((double)2, (double)i);
                }
            }
            return (uint)ret;
        }

        public static BitArray BitArrayCopy(BitArray source, int sourceIndex, int copyCount)
        {
            var ret = new bool[copyCount];

            int j = 0;
            for (int i = sourceIndex; i < (copyCount + sourceIndex); i++)
            {
                ret[j] = source[i];
                j++;
            }

            return new BitArray(ret);
        }

        public static string GetHashDump(byte[] bytes)
        {
            int colCount = 16;
            string ret = string.Format("({0:X04}) ", 0);

            int colCounter = 0;
            for (int i = 0; i < bytes.Length; i++)
            {
                if (colCounter == colCount)
                {
                    colCounter = 0;
                    ret += Environment.NewLine + string.Format("({0:X04}) ", i);
                }
                ret += string.Format("{0:X02} ", bytes[i]);
                colCounter++;
            }

            return ret;
        }

        // Would have been nice to use an extension method here... -- flgr
        public static string StringInspect(string s)
        {
            var result = new StringBuilder("\"");

            foreach (char c in s)
            {
                switch (c)
                {
                    case '\r':
                        result.Append(@"\r");
                        break;

                    case '\n':
                        result.Append(@"\n");
                        break;

                    case '\v':
                        result.Append(@"\v");
                        break;

                    default:
                        if (Char.IsControl(c))
                            result.AppendFormat("\\x{0:X2}", (int)c);
                        else
                            result.Append(c);
                        break;
                }
            }

            result.Append("\"");

            return result.ToString();
        }

        public static string GetWritableString(string s)
        {
            var result = new StringBuilder();
            foreach (char c in s)
            {
                if ((int)c >= 0x20)
                {
                    result.Append(c);
                }
            }
            return result.ToString();
        }
    }
}
