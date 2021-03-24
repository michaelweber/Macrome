using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Macrome.Encryption
{
    public class BaseNEncoder
    {
        public Dictionary<int, char> EncodingLookup;
        public Dictionary<char, int> DecodingLookup;
        public int DictionarySize;

        public static string GetCP1252Dictionary()
        {
            var encoding = Encoding.GetEncoding(1252);
            string dictionary = "";
            for (int i = 1; i < 256; i += 1)
            {
                dictionary += encoding.GetString(new []{(byte)i});
            }

            return dictionary;
        }

        public void BuildLookupsFromDictionaryString(string dictionary)
        {
            EncodingLookup = new Dictionary<int, char>();
            DecodingLookup = new Dictionary<char, int>();
            
            int value = 0;
            foreach (char c in dictionary)
            {
                EncodingLookup.Add(value, c);
                DecodingLookup.Add(c, value);
                value += 1;

            }

            DictionarySize = value;
        }
        
        public BaseNEncoder(
            string encodingDictionary = null)
        {
            if (encodingDictionary == null)
            {
                encodingDictionary = GetCP1252Dictionary();
            }   
            
            BuildLookupsFromDictionaryString(encodingDictionary);
        }

        public string Encode(byte[] bytesToEncode)
        {
            // There's a MUCH cleverer way to do this, but we need to be able to write the decoder in VBA so keeping it
            // simple here even though it's not QUITE as byte-efficient.
            
            // Essentially we just map the byte value to our dictionary unless greater than the "max" value (DictSize - 2)
            // If it's greater than the max value, then we subtract the "max" value from our current value, and append that
            // encoded value along with our "end of encoded value" marker which is stored in DictSize - 1.
            string encodedString = "";
            foreach (byte b in bytesToEncode)
            {
                if (b == DictionarySize - 2)
                {
                    encodedString += EncodingLookup[DictionarySize - 2];
                    encodedString += EncodingLookup[DictionarySize - 1];
                }
                else if (b >= DictionarySize - 1)
                {
                    encodedString += EncodingLookup[DictionarySize - 2];
                    encodedString += EncodingLookup[(int) b - (DictionarySize - 2)];
                    encodedString += EncodingLookup[DictionarySize - 1];
                }
                else
                {
                    encodedString += EncodingLookup[(int) b];
                }
            }
            return encodedString;
        }

        public byte[] Decode(string strToDecode)
        {
            // There are essentially 3 possible cases when we decode a byte:
            // 1. The byte can be directly mapped / described in one byte and is NOT our possible maximum byte value
            // 2. The byte IS exactly our possible single-char maximum value (DictionarySize - 2) - and is followed by a separator value (DictionarySize - 1)
            // 3. The byte IS greater than our single-char max value and is followed by a value to add, then is followed by a separator value.
            
            List<byte> bytes = new List<byte>();
            int offset = 0;
            while (offset < strToDecode.Length)
            {
                char currentChar = strToDecode[offset];
                int lookup = DecodingLookup[currentChar];
                if (lookup == DictionarySize - 2)
                {
                    int nextLookup = DecodingLookup[strToDecode[offset + 1]];
                    if (nextLookup == DictionarySize - 1)
                    {
                        bytes.Add((byte)lookup);
                        offset += 1;
                    }
                    else
                    {
                        bytes.Add((byte)(lookup + nextLookup));
                        offset += 2;
                    }
                }
                else
                {
                    bytes.Add((byte)lookup);
                }
                offset += 1;
            }

            return bytes.ToArray();
        }
    }
}