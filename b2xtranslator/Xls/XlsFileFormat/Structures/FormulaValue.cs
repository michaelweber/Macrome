using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.xls.XlsFileFormat.Structures
{
    public class FormulaValue
    {
        public static FormulaValue GetStringFormulaValue()
        {
            return new FormulaValue();
        }

        public static FormulaValue GetBooleanFormulaValue(bool value)
        {
            return new FormulaValue()
            {
                byte1 = 0x1,
                byte3 = Convert.ToByte(value ? 0x1 : 0x0)
            };
        }

        public static FormulaValue GetEmptyStringFormulaValue()
        {
            return new FormulaValue()
            {
                byte1 = 0x3,
            };
        }

        public static FormulaValue GetDoubleFormulaValue(double xNum)
        {
            byte[] xNumBytes = BitConverter.GetBytes(xNum);
            return new FormulaValue()
            {
                byte1 = xNumBytes[0],
                byte2 = xNumBytes[1],
                byte3 = xNumBytes[2],
                byte4 = xNumBytes[3],
                byte5 = xNumBytes[4],
                byte6 = xNumBytes[5],
                fExprO = new byte[] {xNumBytes[6], xNumBytes[7]}
            };
        }

        private FormulaValue()
        {

        }


        public byte byte1 = 0;
        public byte byte2 = 0;
        public byte byte3 = 0;
        public byte byte4 = 0;
        public byte byte5 = 0; 
        public byte byte6 = 0;
        public byte[] fExprO = {0xFF, 0xFF};

        public byte[] Bytes
        {
            get
            {
                byte[] bytes = 
                {
                    byte1,
                    byte2,
                    byte3,
                    byte4,
                    byte5,
                    byte6,
                    fExprO[0],
                    fExprO[1]
                };
                return bytes;
            }
        }
    }


}
