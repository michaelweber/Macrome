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

        [Test]
        public void TestBase255EncodingForBasicPayload()
        {
            byte[] shellcode = TestHelpers.GetPopCalcShellcode();
            
            BaseNEncoder encoder = new BaseNEncoder();
            string encodedShellcode = encoder.Encode(shellcode);
            int encodedLen = encodedShellcode.Length;
            byte[] decodedShellcode = encoder.Decode(encodedShellcode);
            Assert.AreEqual(shellcode, decodedShellcode);
        }
        
        [Test]
        public void TestBase255EncodingForEdgeCasePayloads()
        {
            byte[] shellcode = new byte[] {0x00, 0xFD, 0xFE, 0xFF};
            
            BaseNEncoder encoder = new BaseNEncoder();
            string encodedShellcode = encoder.Encode(shellcode);
            int encodedLen = encodedShellcode.Length;
            byte[] decodedShellcode = encoder.Decode(encodedShellcode);
            Assert.AreEqual(shellcode, decodedShellcode);
        }
    }
}
