using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using b2xtranslator.OpenXmlLib.SpreadsheetML;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.StructuredStorage.Reader;
using Macrome;
using Macrome.Encryption;
using NUnit.Framework;

namespace UnitTests
{
    [TestFixture]
    public class EncryptionTest
    {

        [SetUp]
        public void SetUp()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        // FilePass Structure for VelvetSweatshop XOR Obfuscation
        // 2F 00 06 00 00 00 59 B3 0A 9A 
        // [RID] [LEN] [TYPE][key][vrfy]
        [Test]
        public void TestPasswordVerifier()
        {
            XorObfuscation xorObfuscation = new XorObfuscation();

            ushort result = xorObfuscation.CreatePasswordVerifier_Method1(XorObfuscation.DefaultPassword);

            //verification bytes = 0x9A0A
            Assert.AreEqual(0x9A0A, result);
        }

        [Test]
        public void TestPasswordXorKeyDerivation()
        {
            XorObfuscation xorObfuscation = new XorObfuscation();

            ushort key = xorObfuscation.CreateXorKey_Method1(XorObfuscation.DefaultPassword);

            //key bytes = 0xB359
            Assert.AreEqual(0xB359, key);
        }


        [Test]
        public void TestXorEncryptionDecryption()
        {
            XorObfuscation xorObfuscation = new XorObfuscation();

            byte[] plaintextBytes = new byte[] {0x30, 0x31, 0x32, 0x33, 0x34, 0x35, 0x36, 0x37, 0x38, 0x39};

            byte[] encryptedBytes =
                xorObfuscation.EncryptData_Method1(XorObfuscation.DefaultPassword, plaintextBytes, 0);

            byte[] decryptedBytes =
                xorObfuscation.DecryptData_Method1(XorObfuscation.DefaultPassword, encryptedBytes, 0);

            for (int i = 0; i < plaintextBytes.Length; i += 1)
            {
                Assert.AreNotEqual(plaintextBytes[i], encryptedBytes[i]);
                Assert.AreEqual(plaintextBytes[i], decryptedBytes[i]);
            }
        }

        [Test]
        public void TestXorEncryptionDecryption2()
        {
            XorObfuscation xorObfuscation = new XorObfuscation();

            byte[] mmsBytes = new byte[] { 0x00, 0x00, 0xB0, 0x04 };
            byte mmsByteOffset = 6;

            byte[] encryptedBytes =
                xorObfuscation.EncryptData_Method1(XorObfuscation.DefaultPassword, mmsBytes, mmsByteOffset);

            byte[] decryptedBytes =
                xorObfuscation.DecryptData_Method1(XorObfuscation.DefaultPassword, encryptedBytes, mmsByteOffset);

            for (int i = 0; i < mmsBytes.Length; i += 1)
            {
                Assert.AreNotEqual(mmsBytes[i], encryptedBytes[i]);
                Assert.AreEqual(mmsBytes[i], decryptedBytes[i]);
            }
        }

        [Test]
        public void TestDecryptObfuscatedDoc()
        {
            WorkbookStream obfuscatedStream = TestHelpers.GetXorObfuscatedWorkbookStream();
            XorObfuscation xorObfuscation = new XorObfuscation();
            WorkbookStream deobfuscatedStream = xorObfuscation.DecryptWorkbookStream(obfuscatedStream);

            foreach (var record in deobfuscatedStream.Records)
            {
                Console.WriteLine(record.ToHexDumpString(0x1000, false));
            }

            List<Lbl> AutoOpenLabels = deobfuscatedStream.GetAutoOpenLabels();
            Assert.AreEqual(1, AutoOpenLabels.Count);

            List<FilePass> filePassRecords = deobfuscatedStream.GetAllRecordsByType<FilePass>();
            Assert.AreEqual(0, filePassRecords.Count);
        }

        [Test]
        public void TestEncryptDoc()
        {
            WorkbookStream obfuscatedStream = TestHelpers.GetXorObfuscatedWorkbookStream();
            XorObfuscation xorObfuscation = new XorObfuscation();
            WorkbookStream deobfuscatedStream = xorObfuscation.DecryptWorkbookStream(obfuscatedStream);
            WorkbookStream reObfuscatedStream = xorObfuscation.EncryptWorkbookStream(deobfuscatedStream);

            foreach (var record in reObfuscatedStream.Records)
            {
                Console.WriteLine(record.ToHexDumpString(0x1000, false));
            }

            CodePage originalCpRecord = obfuscatedStream.GetAllRecordsByType<CodePage>().First();
            CodePage newCpRecord = reObfuscatedStream.GetAllRecordsByType<CodePage>().First();
            Assert.AreEqual(originalCpRecord.cv, newCpRecord.cv);
        }

        [Test]
        public void GenerateFilePassBytesForPasswords()
        {
            XorObfuscation xorObfuscation = new XorObfuscation();

            byte[] xorArray = xorObfuscation.CreateXorArray_Method1(XorObfuscation.DefaultPassword);
            string xorString = BitConverter.ToString(xorArray).Replace("-", " ");
            Console.WriteLine(xorString);

            string[] passwords = new[] { "000", "111", "222", "333", "444", "555", "5 5 5", "666", "777", "7 7 7", "888", "999", "123", "321", "987", "876", "543", "321", "135", "246", "357", "468", "579", "1234", "2020", "2021", "password", "password1", "infected", "infected2" };
            foreach (var password in passwords)
            {
                Console.WriteLine(string.Format("FilePass Bytes for {0}", password));
                FilePass fp  = xorObfuscation.CreateFilePassFromPassword(password);
                Console.WriteLine(fp.ToHexDumpString(16, false, true));
            }
        }

        [Test]
        public void GenerateVelvetSweatshopMacroSheetSigs()
        {
            XorObfuscation xorObfuscation = new XorObfuscation();

            BoundSheet8 sheetVersion1 = new BoundSheet8(BoundSheet8.HiddenState.Visible, BoundSheet8.SheetType.Macrosheet, "Sheet");
            BoundSheet8 sheetVersion2 = new BoundSheet8(BoundSheet8.HiddenState.Hidden, BoundSheet8.SheetType.Macrosheet, "Sheet1");
            BoundSheet8 sheetVersion3 = new BoundSheet8(BoundSheet8.HiddenState.VeryHidden, BoundSheet8.SheetType.Macrosheet, "Sheet2");

            BoundSheet8[] sheets = new[] {sheetVersion1, sheetVersion2, sheetVersion3};

            int hiddenState = 0;
            foreach (BoundSheet8 currentSheet in sheets)
            {
                for (int offset = 0; offset < 16; offset += 1)
                {
                    byte[] header = currentSheet.GetBytes().Take(4).ToArray();
                    byte[] data = currentSheet.GetBytes().Skip(4).ToArray();
                    byte[] encryptedData = xorObfuscation.EncryptData_Method1(XorObfuscation.DefaultPassword, data, (byte)offset);

                    string ruleName = string.Format("$boundsheet8_hs{0}_vs_xor_offset{1}", hiddenState, offset);

                    string yaraSig = ruleName + " = { ";
                    yaraSig += BitConverter.ToString(header.Take(2).ToArray()).Replace("-", " ");
                    yaraSig += " [6] ";
                    yaraSig += BitConverter.ToString(encryptedData.Skip(4).Take(2).ToArray()).Replace("-", " ");
                    yaraSig += " }";
                    Console.WriteLine(yaraSig);
                }

                hiddenState += 1;
            }
        }

    }
}
