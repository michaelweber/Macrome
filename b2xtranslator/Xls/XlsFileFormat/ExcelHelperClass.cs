using System;
using System.Collections.Generic;
using System.Linq;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat
{
    /// <summary>
    /// Assembly of some static methods 
    /// </summary>
    public class ExcelHelperClass
    {
        /// <summary>
        /// This method is used to parse a RK RecordType  
        /// This is a special MS Excel format that is used to store floatingpoint and integer values 
        /// The floatingpoint arithmetic is a little bit different to the IEEE standard format.  
        /// 
        /// Integer and floating point values are booth stored in a RK Record. To differ between them 
        /// it is necessary to check the first two bits from this value. 
        /// </summary>
        /// <param name="rk">Bytestream from the RK Record</param>
        /// <returns>The correct parsed number</returns>
        public static double NumFromRK(Byte[] rk)
        {
            double num = 0;
            int high = 1023;
            uint number;
            number = System.BitConverter.ToUInt32(rk, 0);
            // Select which type of number 
            uint type = number & 0x00000003;
            // if the last two bits are 00 or 01 then it is floating point IEEE number 
            // type 0 and type 1 expects booth the same arithmetic 
            if (type == 0 || type == 1)
            {

                uint mant = 0;
                // masking the mantisse 
                mant = number & 0x000ffffc;
                // shifting the result by 2  
                mant = mant >> 2;

                uint exp = 0;
                // masking the exponent 
                exp = number & 0x7ff00000;
                // shifting the exponent by 20 
                exp = exp >> 20;
                // (1 + (Mantisse / 2^18)) * 2 ^ (Exponent - 1023) 
                num = (1 + (mant / System.Math.Pow(2.0, 18.0))) * System.Math.Pow(2, (double)(exp - high));
                // now there is a sign bit too, the highest bit from the RK Record 
                uint signBit = number & 0x80000000;
                // shifting the value by 31 bit 
                signBit = signBit >> 31;
                // if the signBit is 0 it is a positive number, otherwise it is a negative number 
                if (signBit == 1)
                {
                    num *= -1;
                }
                if (type == 1)
                {
                    num /= 100;

                }
            }
            // if type is 1 then it is an IEEE number * 100 

            // if type is 2 or 3 it is an integer value
            else if (type == 2 || type == 3)
            {
                // 30 bits for the integer value, 2 bits for the type identification 
                uint unumber = 0;
                unumber = number & 0xfffffffc;
                // shifting the value by two 
                unumber = unumber >> 2;
                num = (double)unumber;
                if (type == 3)
                {
                    num /= 100;
                }
            }
            // if type is 3, it has to be multiplicated with 100

            return num;
        }

        /// <summary>
        /// converts the integer column value to a string like AB 
        /// excel binary format has a cap at column 256 -> IV, so there is no need to 
        /// create an almighty algorithm ;) 
        /// </summary>
        /// <returns>String</returns>
        public static string intToABCString(int colnumber, string rownumber)
        {

            string value = "";
            int remain = 0;
            if (colnumber < 26)
            {
                value += (char)(colnumber + 65);
            }
            else if (colnumber < Math.Pow(26, 2))
            {
                remain = colnumber % 26;
                colnumber = colnumber / 26;
                value += (char)(colnumber + 64);
                value = value + (char)(remain + 65);
            }
            else if (colnumber < Math.Pow(26, 3))
            {
                remain = colnumber % (int)Math.Pow(26, 2);
                colnumber = colnumber / (int)Math.Pow(26, 2);
                value += (char)(colnumber + 64);
                colnumber = remain;
                remain = colnumber % 26;
                colnumber = colnumber / 26;
                value = value + (char)(colnumber + 64);
                value = value + (char)(remain + 65);
            }
            return value + rownumber;
        }

        /// <summary>
        /// converts the integer column value to a string like AB 
        /// excel binary format has a cap at column 256 -> IV, so there is no need to 
        /// create an almighty algorithm ;) 
        /// </summary>
        /// <returns>String</returns>
        public static string intToABCString(int colnumber, string rownumber, bool colRelative, bool rwRelative)
        {

            string value = "";
            int remain = 0;
            if (colnumber < 26)
            {
                value += (char)(colnumber + 65);
            }
            else if (colnumber < Math.Pow(26, 2))
            {
                remain = colnumber % 26;
                colnumber = colnumber / 26;
                value += (char)(colnumber + 64);
                value = value + (char)(remain + 65);
            }
            else if (colnumber < Math.Pow(26, 3))
            {
                remain = colnumber % (int)Math.Pow(26, 2);
                colnumber = colnumber / (int)Math.Pow(26, 2);
                value += (char)(colnumber + 64);
                colnumber = remain;
                remain = colnumber % 26;
                colnumber = colnumber / 26;
                value = value + (char)(colnumber + 64);
                value = value + (char)(remain + 65);
            }
            if (!colRelative)
                value = "$" + value;

            if (!rwRelative)
                rownumber = "$" + rownumber;

            return value + rownumber;
        }


        public static Stack<AbstractPtg> getFormulaStack(IStreamReader reader, ushort cce)
        {
            var ptgStack = new Stack<AbstractPtg>();
            try
            {
                for (uint i = 0; i < cce; i++)
                {
                    var ptgtype = (PtgNumber)reader.ReadByte();

                    AbstractPtg.PtgDataType dt = AbstractPtg.PtgDataType.REFERENCE;

                    if ((int)ptgtype > 0x5D)
                    {
                        ptgtype -= 0x40;
                        dt = AbstractPtg.PtgDataType.ARRAY;
                    }
                    
                    else if ((int)ptgtype > 0x3D)
                    {
                        ptgtype -= 0x20;
                        dt = AbstractPtg.PtgDataType.VALUE;
                    }
                    AbstractPtg ptg = null;
                    if (ptgtype == PtgNumber.Ptg0x19Sub)
                    {
                        var ptgtype2 = (Ptg0x19Sub)reader.ReadByte();
                        switch (ptgtype2)
                        {
                            case Ptg0x19Sub.PtgAttrSum: ptg = new PtgAttrSum(reader, ptgtype2); break;
                            case Ptg0x19Sub.PtgAttrIf: ptg = new PtgAttrIf(reader, ptgtype2); break;
                            case Ptg0x19Sub.PtgAttrGoto: ptg = new PtgAttrGoto(reader, ptgtype2); break;
                            case Ptg0x19Sub.PtgAttrSemi: ptg = new PtgAttrSemi(reader, ptgtype2); break;
                            case Ptg0x19Sub.PtgAttrChoose: ptg = new PtgAttrChoose(reader, ptgtype2); break;
                            case Ptg0x19Sub.PtgAttrSpace: ptg = new PtgAttrSpace(reader, ptgtype2); break;
                            case Ptg0x19Sub.PtgAttrBaxcel1: ptg = new PtgAttrBaxcel(reader, ptgtype2, false); break;
                            case Ptg0x19Sub.PtgAttrBaxcel2: ptg = new PtgAttrBaxcel(reader, ptgtype2, true); break;
                            case Ptg0x19Sub.PtgNotDocumented: ptg = new PtgNotDocumented(reader, ptgtype2); break;
                            default: break;
                        }
                    }
                    else if (ptgtype == PtgNumber.Ptg0x18Sub)
                    {

                    }
                    else
                    {
                        switch (ptgtype)
                        {
                            case PtgNumber.PtgInt: ptg = new PtgInt(reader, ptgtype); break;
                            case PtgNumber.PtgAdd: ptg = new PtgAdd(reader, ptgtype); break;
                            case PtgNumber.PtgSub: ptg = new PtgSub(reader, ptgtype); break;
                            case PtgNumber.PtgMul: ptg = new PtgMul(reader, ptgtype); break;
                            case PtgNumber.PtgDiv: ptg = new PtgDiv(reader, ptgtype); break;
                            case PtgNumber.PtgParen: ptg = new PtgParen(reader, ptgtype); break;
                            case PtgNumber.PtgNum: ptg = new PtgNum(reader, ptgtype); break;
                            case PtgNumber.PtgArray: ptg = new PtgArray(reader, ptgtype); break;
                            case PtgNumber.PtgRef: ptg = new PtgRef(reader, ptgtype); break;
                            case PtgNumber.PtgRefN: ptg = new PtgRefN(reader, ptgtype); break;
                            case PtgNumber.PtgPower: ptg = new PtgPower(reader, ptgtype); break;
                            case PtgNumber.PtgPercent: ptg = new PtgPercent(reader, ptgtype); break;
                            case PtgNumber.PtgBool: ptg = new PtgBool(reader, ptgtype); break;
                            case PtgNumber.PtgGt: ptg = new PtgGt(reader, ptgtype); break;
                            case PtgNumber.PtgGe: ptg = new PtgGe(reader, ptgtype); break;
                            case PtgNumber.PtgLt: ptg = new PtgLt(reader, ptgtype); break;
                            case PtgNumber.PtgLe: ptg = new PtgLe(reader, ptgtype); break;
                            case PtgNumber.PtgEq: ptg = new PtgEq(reader, ptgtype); break;
                            case PtgNumber.PtgNe: ptg = new PtgNe(reader, ptgtype); break;
                            case PtgNumber.PtgUminus: ptg = new PtgUminus(reader, ptgtype); break;
                            case PtgNumber.PtgUplus: ptg = new PtgUplus(reader, ptgtype); break;
                            case PtgNumber.PtgStr: ptg = new PtgStr(reader, ptgtype); break;
                            case PtgNumber.PtgConcat: ptg = new PtgConcat(reader, ptgtype); break;
                            case PtgNumber.PtgUnion: ptg = new PtgUnion(reader, ptgtype); break;
                            case PtgNumber.PtgIsect: ptg = new PtgIsect(reader, ptgtype); break;
                            case PtgNumber.PtgMemErr: ptg = new PtgMemErr(reader, ptgtype); break;
                            case PtgNumber.PtgArea: ptg = new PtgArea(reader, ptgtype); break;
                            case PtgNumber.PtgAreaN: ptg = new PtgAreaN(reader, ptgtype); break;
                            case PtgNumber.PtgFuncVar: ptg = new PtgFuncVar(reader, ptgtype); break;
                            case PtgNumber.PtgFunc: ptg = new PtgFunc(reader, ptgtype); break;
                            case PtgNumber.PtgExp: ptg = new PtgExp(reader, ptgtype); break;
                            case PtgNumber.PtgRef3d: ptg = new PtgRef3d(reader, ptgtype); break;
                            case PtgNumber.PtgArea3d: ptg = new PtgArea3d(reader, ptgtype); break;
                            case PtgNumber.PtgNameX: ptg = new PtgNameX(reader, ptgtype); break;
                            case PtgNumber.PtgName: ptg = new PtgName(reader, ptgtype); break;
                            case PtgNumber.PtgMissArg: ptg = new PtgMissArg(reader, ptgtype); break;
                            case PtgNumber.PtgRefErr: ptg = new PtgRefErr(reader, ptgtype); break;
                            case PtgNumber.PtgRefErr3d: ptg = new PtgRefErr3d(reader, ptgtype); break;
                            case PtgNumber.PtgAreaErr: ptg = new PtgAreaErr(reader, ptgtype); break;
                            case PtgNumber.PtgAreaErr3d: ptg = new PtgAreaErr3d(reader, ptgtype); break;
                            case PtgNumber.PtgMemFunc: ptg = new PtgMemFunc(reader, ptgtype); break;
                            case PtgNumber.PtgMemArea: ptg = new PtgMemArea(reader, ptgtype); break;
                            case PtgNumber.PtgErr: ptg = new PtgErr(reader, ptgtype); break;
                            case PtgNumber.PtgRange: ptg = new PtgRange(reader, ptgtype); break;

                            default: break;
                        }
                    }
                    i += ptg.getLength() - 1;

                    ptg.dataType = dt;
                    ptgStack.Push(ptg);
                }
            }
            catch (Exception ex)
            {
                throw new ExtractorException(ExtractorException.PARSEDFORMULAEXCEPTION, ex);
            }

            return ptgStack;
        }

        public static string parseVirtualPath(string path)
        {
            // NOTE: A virtual path must be a string in the following grammar: 
            //
            //    virt-path = volume / unc-volume / rel-volume / transfer-protocol / startup / alt-startup / library / simple-file-path / ole-link
            path = path.Trim();

            if (path.StartsWith("\x0001\x0001\x0040"))
            {
                // unc-volume = %x0001 %x0001 %x0040 unc-path
                path = path.Substring(3);
            }
            else if (path.StartsWith("\x0001\x0001"))
            {
                // volume     = %x0001 %x0001 volume-character file-path
                // path[2] is a volumn character in the range %x0041-%x005A / %x0061-%x007A
                path = path.Substring(2, 1) + ":\\" + path.Substring(3);
            }
            else if (path.StartsWith("\x0001\x0002"))
            {
                // rel-volume = %x0001 %x0002 file-path
                path = path.Substring(2);
            }
            else if (path.StartsWith("\x0001\x0005"))
            {
                // transfer-protocol = %x0001 %x0005 count transfer-path
                // count is ignored 
                path = path.Substring(3);
            }
            else if (path.StartsWith("\x0001\x0006"))
            {
                // startup = %x0001 %x0006 file-path
                // TODO: map startup path
                path = path.Substring(2);
            }
            else if (path.StartsWith("\x0001\x0007"))
            {
                // alt-startup = %x0001 %x0007 file-path
                path = path.Substring(2);
            }
            else if (path.StartsWith("\x0001\x0008"))
            {
                // library = %x0001 %x0008 file-path
                // TODO: map library path
                path = "file:///" + path.Substring(2);
            }
            else if (path.StartsWith("\x0001"))
            {
                // simple-file-path = [%x0001] file-path
                //   (\x0001 is optional, but if it is missing path is
                //    already set to the correct value)
                path = path.Substring(1);
            }
            

            /// Replace 0x03 with \
            path = path.Replace((char)0x03, '\\');
            /// replace ' ' with %20
            ///path = path.Replace(" ", "%20");


            return path;
        }

        /// <summary>
        /// This method reads x bytes from a IStreamReader to get a string from this
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="cch"></param>
        /// <param name="grbit"></param>
        /// <returns></returns>
        public static string getStringFromBiffRecord(IStreamReader reader, int cch, int grbit)
        {
            string value = "";
            if (grbit == 0)
            {
                for (int i = 0; i < cch; i++)
                {
                    value += (char)reader.ReadByte();
                }
            }
            else
            {
                for (int i = 0; i < cch; i++)
                {
                    value += System.BitConverter.ToChar(reader.ReadBytes(2), 0);
                }
            }
            return value;
        }

        /// <summary>
        /// Converts the builtin function id to a string
        /// </summary>
        /// <param name="idstring"></param>
        /// <returns></returns>
        public static string getNameStringfromBuiltInFunctionID(string idstring)
        {
            char firstChar = (char)idstring.ToCharArray().GetValue(0);

            switch (firstChar)
            {
                case (char)0x00: return "Consolidate_Area";
                case (char)0x01: return "Auto_Open";
                case (char)0x02: return "Auto_Close";
                case (char)0x03: return "Extract";
                case (char)0x04: return "Database";
                case (char)0x05: return "Criteria";
                case (char)0x06: return "Print_Area";
                case (char)0x07: return "Print_Titles";
                case (char)0x08: return "Recorder";
                case (char)0x09: return "Data_Form";
                case (char)0x0A: return "Auto_Activate";
                case (char)0x0B: return "Auto_Deactivate";
                case (char)0x0C: return "Sheet_Title";
                case (char)0x0D: return "_FilterDatabase";


                default: return idstring;

            }
        }

        /// <summary>
        /// This method reads x bytes from a IStreamReader to get a hyperlink string from this
        /// </summary>
        /// <param name="reader"></param>
        /// <param name="cch"></param>
        /// <param name="grbit"></param>
        /// <returns></returns>
        public static string getHyperlinkStringFromBiffRecord(IStreamReader reader)
        {
            string value = "";
            uint length = reader.ReadUInt32();
            for (int i = 0; i < length; i++)
            {
                value += System.BitConverter.ToChar(reader.ReadBytes(2), 0);
            }


            return value.Remove(value.Length - 1); ;
        }

        /// <summary>
        /// Escapes special characters in strings so that they can be safely used in formulas.
        /// </summary>
        /// <param name="unescapedString">The input string to be escaped.</param>
        /// <remarks>This method currently escapes double and single quotes.</remarks>
        public static string EscapeFormulaString(string unescapedString)
        {
            return unescapedString.Replace("\"", "\"\"")
                .Replace("'", "''");
        }

        public static string UnescapeFormulaString(string escapedString)
        {
            return escapedString.Replace("\"\"","\"" )
                .Replace("''","'");

        }

        private static int GetColNumberFromExcelA1ColName(string colName)
        {
            if (colName.Length == 1)
            {
                return (int)Convert.ToByte(colName[0]) - (int)'A' + 1;
            }
            else
            {
                return ((int)Convert.ToByte(colName[0]) - (int)'A' + 1) * 26 +
                       ((int)Convert.ToByte(colName[1]) - (int)'A' + 1);
            }
        }
        public static string ConvertA1ToR1C1(string a1)
        {
            string colString = new string(a1.TakeWhile(c => !Char.IsNumber(c)).ToArray());
            int rowNum = Int32.Parse(new string(a1.TakeLast(a1.Length - colString.Length).ToArray()));
            int colNum = GetColNumberFromExcelA1ColName(colString);
            return string.Format("R{0}C{1}", rowNum, colNum);
        }

        public static string GetExcelA1ColNameFromColNumber(int colNumber)
        {
            if (colNumber <= 26)
            {
                char firstColLetter = (char)(colNumber + (int)'A' - 1);
                return "" + firstColLetter;
            }
            else
            {
                char firstColLetter = (char)(colNumber % 26 + (int)'A' - 1);
                char secondColLetter = (char)((int)(colNumber / 26) + (int)'A' - 1);

                //There's an edge case where if colNumber % 26 is 0, we increment both letters one character too far
                if (colNumber % 26 == 0)
                {
                    firstColLetter = 'Z';
                    secondColLetter = (char)((int)(colNumber / 26) + (int)'A' - 2);
                }


                return "" + secondColLetter + firstColLetter;
            }
        }

        public static string ConvertR1C1ToA1(string r1c1)
        {
            string rowNumString = new string(r1c1.Skip(1).TakeWhile(c => Char.IsNumber(c)).ToArray());
            int rowNum = Convert.ToInt32(rowNumString);
            string colNumString = new string(r1c1.SkipWhile(c => c != 'C' && c != 'c').Skip(1).ToArray());
            int colNum = Convert.ToInt32(colNumString);
            string a1String = string.Format("{0}{1}", GetExcelA1ColNameFromColNumber(colNum), rowNum);
            return a1String;
        }
    }
}
