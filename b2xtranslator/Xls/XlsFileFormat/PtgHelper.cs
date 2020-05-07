using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.Tools;
using b2xtranslator.xls.XlsFileFormat.Structures;

namespace b2xtranslator.xls.XlsFileFormat
{
    public static class PtgHelper
    {
        public static Type[] Byte19PtgTypes = new[]
        {
            typeof(PtgAttrSemi),
            typeof(PtgAttrIf),
            typeof(PtgAttrChoose),
            typeof(PtgAttrGoto),
            typeof(PtgAttrSum),
            typeof(PtgAttrSpace),
            typeof(PtgNotDocumented)
        };

        private static byte[] GetByteDataFromPtg(AbstractPtg ptg)
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);
            switch (ptg)
            {
                case PtgInt ptgInt: 
                    bw.Write(Convert.ToUInt16(ptg.getData()));
                    break;
                case PtgAdd ptgAdd:
                case PtgSub ptgSub:
                case PtgMul ptgMul:
                case PtgDiv ptgDiv:
                case PtgParen ptgParen:
                case PtgPower ptgPower:
                case PtgPercent ptgPercent:
                case PtgGt ptgGt:
                case PtgGe ptgGe:
                case PtgLt ptgLt:
                case PtgLe ptgLe:
                case PtgEq ptgEq:
                case PtgNe ptgNe:
                case PtgUminus ptgUminus:
                case PtgUplus ptgUplus:
                case PtgConcat ptgConcat:
                case PtgUnion ptgUnion:
                case PtgIsect ptgIsect:
                    //No Data
                    break;
                case PtgNum ptgNum:
                    bw.Write(Convert.ToDouble(ptg.getData(), CultureInfo.GetCultureInfo("en-US")));
                    break;
                case PtgRef ptgRef:
                    RgceLoc loc = new RgceLoc(ptgRef.rw, ptgRef.col, ptgRef.colRelative, ptgRef.rwRelative);
                    bw.Write(loc.Bytes);
                    break;
                case PtgRefN ptgRefN:
                    RgceLoc loc2 = new RgceLoc(Convert.ToUInt16(ptgRefN.rw), Convert.ToUInt16(ptgRefN.col), ptgRefN.colRelative, ptgRefN.rwRelative);
                    bw.Write(loc2.Bytes);
                    break;
                case PtgBool ptgBool:
                    string boolData = ptgBool.getData();
                    bw.Write((boolData.Equals("TRUE") ? (byte)1 : (byte)0));
                    break;
                case PtgStr ptgStr:
                    string unescapedString = ExcelHelperClass.UnescapeFormulaString(ptgStr.getData());
                    ShortXLUnicodeString ptgString = new ShortXLUnicodeString(unescapedString);
                    bw.Write(ptgString.Bytes);
                    break;
                case PtgFuncVar ptgFuncVar:
                    bw.Write(ptgFuncVar.cparams);
                    ushort tab = ptgFuncVar.tab;
                    ushort fCeFunc = ptgFuncVar.fCelFunc ? (ushort)0x8000 : (ushort)0;
                    tab = (ushort)(tab | fCeFunc);
                    bw.Write(tab);
                    break;
                case PtgFunc ptgFunc:
                    bw.Write(Convert.ToUInt16(ptgFunc.tab));
                    break;
                case PtgExp ptgExp:
                    RgceLoc loc3 = new RgceLoc(Convert.ToUInt16(ptgExp.rw), Convert.ToUInt16(ptgExp.col), false, false);
                    bw.Write(loc3.Bytes);
                    break;
                case PtgRef3d ptgRef3d:
                    RgceLoc loc4 = new RgceLoc(Convert.ToUInt16(ptgRef3d.rw), Convert.ToUInt16(ptgRef3d.col), ptgRef3d.colRelative, ptgRef3d.rwRelative);
                    bw.Write(Convert.ToUInt16(ptgRef3d.ixti));
                    bw.Write(loc4.Bytes);
                    break;
                case PtgMemErr ptgMemErr:
                case PtgArea ptgArea:
                case PtgAreaN ptgAreaN:
                case PtgArea3d ptgArea3d:
                case PtgNameX ptgNameX:
                case PtgName ptgName:
                case PtgMissArg ptgMissArg:
                case PtgRefErr ptgRefErr:
                case PtgRefErr3d ptgRefErr3d:
                case PtgAreaErr ptgAreaErr:
                case PtgAreaErr3d ptgAreaErr3d:
                case PtgMemFunc ptgMemFunc:
                case PtgErr ptgErr:
                case PtgAttrSum ptgAttrSum: //Start 0x19 ## Section
                case PtgAttrIf ptgAttrIf:
                case PtgAttrGoto ptgAttrGoto:
                case PtgAttrSemi ptgAttrSemi:
                case PtgAttrChoose ptgAttrChoose:
                case PtgAttrSpace ptgAttrSpace:
                case PtgNotDocumented ptgNotDocumented:
                default:
                    throw new NotImplementedException(string.Format("No byte conversion implemented for {0}", ptg));
            }

            return bw.GetBytesWritten();
        }

        public static byte[] GetBytes(Stack<AbstractPtg> ptgStack)
        {
            //Reverse the order here since this is normally a stack
            List<AbstractPtg> ptgList = ptgStack.Reverse().ToList();

            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            foreach (var ptg in ptgList)
            {
                bool isByte18PtgHeader = false;
                bool isByte19PtgHeader = false;
                
                isByte19PtgHeader = Byte19PtgTypes.Any(t => t == ptg.GetType());

                PtgNumber ptgNumber = (PtgNumber) ptg.Id;

                if (isByte19PtgHeader) bw.Write(Convert.ToByte(0x19));
                else if (isByte18PtgHeader) bw.Write(Convert.ToByte(0x18));
                else if ((int)ptgNumber >= 0x20)
                {
                    if (ptg.dataType == AbstractPtg.PtgDataType.VALUE) ptgNumber += 0x20;
                    if (ptg.dataType == AbstractPtg.PtgDataType.ARRAY) ptgNumber += 0x40;
                }

                bw.Write(Convert.ToByte(ptgNumber));
                bw.Write(GetByteDataFromPtg(ptg));
            }

            return bw.GetBytesWritten();
        }
    }
}
