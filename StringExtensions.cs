using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Macrome
{
    public static class StringExtensions
    {
        public static List<String> SplitStringIntoChunks(this string str, int maxChunkSize)
        {
            List<string> chunks = new List<string>();

            for (int offset = 0; offset < str.Length; offset += maxChunkSize)
            {
                int bytesRemaining = str.Length - offset;
                int chunkSize = maxChunkSize;
                if (chunkSize > bytesRemaining)
                {
                    chunkSize = bytesRemaining;
                }

                chunks.Add(new string(str.Skip(offset).Take(chunkSize).ToArray()));
            }

            return chunks;
        }
    }
}
