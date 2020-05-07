using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Macrome;
using NUnit.Framework;

namespace UnitTests
{
    [TestFixture]
    class PayloadGenerationTests
    {
        [Test]
        public void TestGenerateMacroForPayload()
        {
            byte[] shellcode = TestHelpers.GetPopCalcShellcode();
            List<string> macros = FormulaHelper.BuildPayloadMacros(shellcode);

            Assert.AreEqual("END", macros.Last());
        }
    }
}
