using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using b2xtranslator.StructuredStorage.Reader;
using Macrome;

namespace UnitTests
{
    public static class TestHelpers
    {
        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);
                return Path.GetDirectoryName(path);
            }
        }

        public static byte[] GetPopCalc64Shellcode()
        {
            return Convert.FromBase64String(
                "SDHJSIHp3f///0iNBe////9Iu6lO4vZjWPLUSDFYJ0gt+P///+L0VQZhEpOwMtSpTqOnIgighf8G0yQGEHmGyQZppHsQeYaJBmmEMxD9Y+MEr8eqEMMUBXKDimF00pVoh++3YpkQOfsPs77oCtJf63Kq97PTclypTuK+5piGs+FPMqboEOqQIg7Cv2KIEYLhsSu36Gx6nKiYr8eqEMMUBQ8jP24Z8xWRrpcHL1u+8KEL2ycWgKqQIg7Gv2KIlJUiQqqy6Bjunaieo31n0LrVeQ+6tzsGq47oFqOvIgK6V0Vuo6ScuKqV8BSqfXGxpStWsb++2Vny1KlO4vZjEH9ZqE/i9iLiw1/GyR0j2Ljv/qMPWFD25W8rfAZhMktk9KijzhkWFl1Jk7o8jZxjAbNdc7E3lQI0kfrMNof2Y1jy1A==");
        }

        public static byte[] GetPopCalcShellcode()
        {
            byte[] buf = new byte[448] {
0x89,0xe7,0xda,0xc8,0xd9,0x77,0xf4,0x5d,0x55,0x59,0x49,0x49,0x49,0x49,0x49,
0x49,0x49,0x49,0x49,0x49,0x43,0x43,0x43,0x43,0x43,0x43,0x37,0x51,0x5a,0x6a,
0x41,0x58,0x50,0x30,0x41,0x30,0x41,0x6b,0x41,0x41,0x51,0x32,0x41,0x42,0x32,
0x42,0x42,0x30,0x42,0x42,0x41,0x42,0x58,0x50,0x38,0x41,0x42,0x75,0x4a,0x49,
0x4b,0x4c,0x48,0x68,0x4d,0x52,0x73,0x30,0x57,0x70,0x55,0x50,0x43,0x50,0x4b,
0x39,0x7a,0x45,0x54,0x71,0x4f,0x30,0x53,0x54,0x6e,0x6b,0x46,0x30,0x54,0x70,
0x4e,0x6b,0x72,0x72,0x54,0x4c,0x4e,0x6b,0x31,0x42,0x36,0x74,0x4e,0x6b,0x30,
0x72,0x65,0x78,0x76,0x6f,0x38,0x37,0x70,0x4a,0x57,0x56,0x55,0x61,0x59,0x6f,
0x6e,0x4c,0x67,0x4c,0x75,0x31,0x63,0x4c,0x74,0x42,0x34,0x6c,0x51,0x30,0x5a,
0x61,0x48,0x4f,0x64,0x4d,0x56,0x61,0x4a,0x67,0x49,0x72,0x4b,0x42,0x71,0x42,
0x70,0x57,0x6c,0x4b,0x53,0x62,0x34,0x50,0x6c,0x4b,0x72,0x6a,0x77,0x4c,0x6c,
0x4b,0x30,0x4c,0x72,0x31,0x50,0x78,0x58,0x63,0x43,0x78,0x76,0x61,0x6e,0x31,
0x42,0x71,0x4e,0x6b,0x56,0x39,0x31,0x30,0x65,0x51,0x7a,0x73,0x4e,0x6b,0x31,
0x59,0x57,0x68,0x49,0x73,0x74,0x7a,0x67,0x39,0x4c,0x4b,0x34,0x74,0x4e,0x6b,
0x45,0x51,0x7a,0x76,0x36,0x51,0x4b,0x4f,0x6c,0x6c,0x39,0x51,0x4a,0x6f,0x74,
0x4d,0x63,0x31,0x78,0x47,0x55,0x68,0x6d,0x30,0x51,0x65,0x68,0x76,0x45,0x53,
0x61,0x6d,0x6c,0x38,0x67,0x4b,0x63,0x4d,0x76,0x44,0x71,0x65,0x4d,0x34,0x42,
0x78,0x6c,0x4b,0x32,0x78,0x34,0x64,0x43,0x31,0x49,0x43,0x63,0x56,0x4c,0x4b,
0x66,0x6c,0x72,0x6b,0x6c,0x4b,0x50,0x58,0x77,0x6c,0x36,0x61,0x4e,0x33,0x6e,
0x6b,0x35,0x54,0x4e,0x6b,0x77,0x71,0x58,0x50,0x6c,0x49,0x30,0x44,0x55,0x74,
0x75,0x74,0x51,0x4b,0x61,0x4b,0x53,0x51,0x31,0x49,0x42,0x7a,0x76,0x31,0x6b,
0x4f,0x6d,0x30,0x73,0x6f,0x61,0x4f,0x62,0x7a,0x6e,0x6b,0x44,0x52,0x78,0x6b,
0x4c,0x4d,0x73,0x6d,0x33,0x5a,0x33,0x31,0x4e,0x6d,0x4e,0x65,0x78,0x32,0x67,
0x70,0x77,0x70,0x35,0x50,0x36,0x30,0x75,0x38,0x45,0x61,0x4e,0x6b,0x30,0x6f,
0x6f,0x77,0x79,0x6f,0x4e,0x35,0x4f,0x4b,0x69,0x70,0x67,0x6d,0x37,0x5a,0x76,
0x6a,0x52,0x48,0x4e,0x46,0x6d,0x45,0x4d,0x6d,0x6f,0x6d,0x69,0x6f,0x4a,0x75,
0x77,0x4c,0x34,0x46,0x33,0x4c,0x36,0x6a,0x6d,0x50,0x79,0x6b,0x49,0x70,0x73,
0x45,0x34,0x45,0x4d,0x6b,0x33,0x77,0x66,0x73,0x61,0x62,0x32,0x4f,0x43,0x5a,
0x77,0x70,0x36,0x33,0x4b,0x4f,0x38,0x55,0x62,0x43,0x55,0x31,0x62,0x4c,0x43,
0x53,0x54,0x6e,0x35,0x35,0x62,0x58,0x50,0x65,0x33,0x30,0x41,0x41 };
            return buf;
        }

        public static byte[] GetTemplateMacroBytes()
        {
            string template = AssemblyDirectory + Path.DirectorySeparatorChar + @"TestDocs\TemplateMacro.xls";
            using (var fs = new FileStream(template, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                return wbBytes;
            }
        }

        public static byte[] GetAutoOpenTestBytes()
        {
            string template = AssemblyDirectory + Path.DirectorySeparatorChar + @"TestDocs\auto_open_test.xls";
            using (var fs = new FileStream(template, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                return wbBytes;
            }
        }

        public static WorkbookStream GetMassSelectUDFArgumentSheet()
        {
            string template = AssemblyDirectory + Path.DirectorySeparatorChar + @"TestDocs\bug_report_1_mass_select_argument.xls";
            using (var fs = new FileStream(template, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                return new WorkbookStream(wbBytes);
            }
        }


        public static WorkbookStream GetBuiltinHiddenLblSheet()
        {
            string template = AssemblyDirectory + Path.DirectorySeparatorChar + @"TestDocs\builtin_hidden_lbl.xls";
            using (var fs = new FileStream(template, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                return new WorkbookStream(wbBytes);
            }
        }

        public static WorkbookStream GetXorObfuscatedWorkbookStream()
        {
            string template = AssemblyDirectory + Path.DirectorySeparatorChar + @"TestDocs\xorobfuscated_doc.xls";
            using (var fs = new FileStream(template, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                return new WorkbookStream(wbBytes);
            }
        }

        public static WorkbookStream GetDefaultMacroTemplate()
        {
            string template = AssemblyDirectory + Path.DirectorySeparatorChar + @"TestDocs\default_macro_template.xls";
            using (var fs = new FileStream(template, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                return new WorkbookStream(wbBytes);
            }
        }

        public static WorkbookStream GetCharRoundMacroWorkbookStream()
        {
            string template = AssemblyDirectory + Path.DirectorySeparatorChar + @"TestDocs\char-round-macro.xls";
            using (var fs = new FileStream(template, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                return new WorkbookStream(wbBytes);
            }
        }
        public static WorkbookStream GetMultiSheetMacroBytes()
        {
            string template = AssemblyDirectory + Path.DirectorySeparatorChar + @"TestDocs\combined-with-auto-open.xls";
            using (var fs = new FileStream(template, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                return new WorkbookStream(wbBytes);
            }

        }

        public static byte[] GetMacroTestBytes()
        {
            string template = AssemblyDirectory + Path.DirectorySeparatorChar + @"TestDocs\macro-test.xls";
            using (var fs = new FileStream(template, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                return wbBytes;
            }
        }

        public static WorkbookStream GetMacroLoopWorkbookStream()
        {
            string template = AssemblyDirectory + Path.DirectorySeparatorChar + @"TestDocs\macro-loop.xls";
            using (var fs = new FileStream(template, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);
                return new WorkbookStream(wbBytes);
            }
        }
    }
}
