using System;
using System.Globalization;

namespace b2xtranslator.StructuredStorage.Common
{

    /// <summary>
    /// Provides methods for masking/unmasking strings in a path
    /// Author: math
    /// </summary>
    static internal class MaskingHandler
    {
        static readonly uint[] CharsToMask = { '%', '\\' };
        

        /// <summary>
        /// Masks the given string
        /// </summary>
        internal static string Mask(string text)
        {
            string result = text;
            foreach (uint character in CharsToMask)
	        {
                result = result.Replace(new string((char)character,1), string.Format(CultureInfo.InvariantCulture, "%{0:X4}", character));
	        }
            return result;
        }


        /// <summary>
        /// Unmasks the given string
        /// </summary>
        internal static string UnMask(string text)
        {
            string result = text;
            foreach (uint character in CharsToMask)
            {
                result = result.Replace(string.Format("%{0:X4}", character), new string((char)character, 1));
            }
            return result;
        }

    }
}
