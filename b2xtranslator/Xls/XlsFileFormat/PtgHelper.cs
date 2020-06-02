using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text.RegularExpressions;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.Tools;
using b2xtranslator.xls.XlsFileFormat.Records;
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
            typeof(PtgNotDocumented),
            typeof(PtgAttrBaxcel)
        };

        public static List<BiffRecord> UpdatePtgNameRecords(List<BiffRecord> records)
        {
            List<Lbl> labelRecords = records.Where(r => r.Id == RecordType.Lbl).Select(r => r.AsRecordType<Lbl>()).ToList();

            List<BiffRecord> updatedRecords = new List<BiffRecord>();
            foreach (var record in records)
            {
                if (record.Id != RecordType.Formula)
                {
                    updatedRecords.Add(record);
                    continue;
                }

                Formula formulaRecord = record.AsRecordType<Formula>();
                List<AbstractPtg> modifiedStack = new List<AbstractPtg>();
                foreach (var ptg in formulaRecord.ptgStack)
                {
                    if (ptg is PtgName)
                    {
                        int index = (ptg as PtgName).nameindex;
                        string matchingLabel = labelRecords[index - 1].Name.Value;
                        modifiedStack.Add(new PtgName(index, matchingLabel));
                    }
                    else
                    {
                        modifiedStack.Add(ptg);
                    }
                }

                modifiedStack.Reverse();

                formulaRecord.SetCellParsedFormula(new CellParsedFormula(new Stack<AbstractPtg>(modifiedStack)));
                updatedRecords.Add(formulaRecord);
            }
            return updatedRecords;
        }

        private static byte[] GetByteDataFromPtg(AbstractPtg ptg)
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);
            switch (ptg)
            {
                case PtgInt ptgInt: 
                    bw.Write(Convert.ToUInt16(ptg.getData()));
                    break;
                case PtgArray ptgArray:
                    bw.Write(ptgArray.DataBytes);
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
                case PtgMissArg ptgMissArg:
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
                    ShortXLUnicodeString ptgString = new ShortXLUnicodeString(unescapedString, ptgStr.isUnicode);
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
                case PtgAttrBaxcel ptgAttrBaxcel:
                    bw.Write(ptgAttrBaxcel.DataBytes);
                    break;
                case PtgAttrIf ptgAttrIf:
                    bw.Write(ptgAttrIf.Offset);
                    break;
                case PtgAttrGoto ptgAttrGoto:
                    bw.Write(ptgAttrGoto.Offset);
                    break;
                case PtgName ptgName:
                    bw.Write(ptgName.nameindex);
                    break;
                case PtgArea ptgArea:
                    bw.Write(ptgArea.rwFirst);
                    bw.Write(ptgArea.rwLast);
                    bw.Write(ptgArea.colFirst);
                    bw.Write(ptgArea.colLast);
                    break;
                case PtgAttrSum ptgAttrSum: //Start 0x19 ## Section
                case PtgAttrSemi ptgAttrSemi:
                case PtgAttrChoose ptgAttrChoose:
                case PtgAttrSpace ptgAttrSpace:
                case PtgNotDocumented ptgNotDocumented:
                case PtgMemErr ptgMemErr:
                
                case PtgAreaN ptgAreaN:
                case PtgArea3d ptgArea3d:
                case PtgNameX ptgNameX:
                case PtgRefErr ptgRefErr:
                case PtgRefErr3d ptgRefErr3d:
                case PtgAreaErr ptgAreaErr:
                case PtgAreaErr3d ptgAreaErr3d:
                case PtgMemFunc ptgMemFunc:
                case PtgErr ptgErr:
                default:
                    throw new NotImplementedException(string.Format("No byte conversion implemented for {0}", ptg));
            }

            return bw.GetBytesWritten();
        }

        private static List<string> GetFormulaStringFunctionArgs(ref Stack<AbstractPtg> ptgStack, int numArgs, bool showAttributes = false)
        {
            Stack<string> args = new Stack<string>();
            for (int param = 0; param < numArgs; param += 1)
            {
                args.Push(GetFormulaStringInner(ref ptgStack, 1, showAttributes));
            }

            return args.ToList();
        }

        private static string GetFormulaStringInner(ref Stack<AbstractPtg> ptgStack, uint argsToPop, bool showAttributes = false)
        {

            for (int i = 0; i < argsToPop; i += 1)
            {
                AbstractPtg nextPtg = ptgStack.Pop();

                if (nextPtg.OpType() == PtgType.Operator)
                {
                    uint popSize = nextPtg.PopSize();
                    switch (nextPtg)
                    {
                        case PtgFunc ptgFunc:
                            int numParams = PtgFunc.GetParamCountFromFtabValue(ptgFunc.Ftab);
                            string formulaString = ptgFunc.Ftab.ToString().Replace('_', '.') + "(";
                            List<string> args = GetFormulaStringFunctionArgs(ref ptgStack, numParams, showAttributes);
                            formulaString += string.Join(",", args);
                            formulaString += ")";
                            return formulaString;
                        case PtgFuncVar ptgFuncVar:
                            string formulaVariableString = "";
                            if (ptgFuncVar.fCelFunc) formulaVariableString += ptgFuncVar.Cetab.ToString().Replace('_','.');
                            else formulaVariableString += ptgFuncVar.Ftab.ToString().Replace('_', '.');

                            //Handle special case for user defined function invocation
                            if (formulaVariableString.Equals(FtabValues.USERDEFINEDFUNCTION.ToString()))
                            {
                                //Pop the args off first, then we can retrieve the function name
                                List<string> varArgs = GetFormulaStringFunctionArgs(ref ptgStack, ptgFuncVar.cparams - 1, showAttributes);
                                //This should be a PtgName entry
                                formulaVariableString = GetFormulaStringInner(ref ptgStack, 1, showAttributes);
                                formulaVariableString += "(";
                                formulaVariableString += string.Join(",", varArgs);
                                formulaVariableString += ")";
                            }
                            else
                            {
                                formulaVariableString += "(";
                                List<string> varArgs = GetFormulaStringFunctionArgs(ref ptgStack, ptgFuncVar.cparams, showAttributes);
                                formulaVariableString += string.Join(",", varArgs);
                                formulaVariableString += ")";
                            }
                            return formulaVariableString;
                        case PtgAttrBaxcel attrBaxcel:
                            if (showAttributes)
                            {
                                return string.Format("[AttrBaxcel:{0}]", attrBaxcel.DataBytes) + GetFormulaStringInner(ref ptgStack, 1, showAttributes);
                            }
                            else
                            {
                                return GetFormulaStringInner(ref ptgStack, 1, showAttributes);
                            }
                        case PtgAttrGoto attrGoto:
                            if (showAttributes)
                            {
                                return string.Format("[AttrGoto:{0}]", attrGoto.Offset) + GetFormulaStringInner(ref ptgStack, 1, showAttributes);
                            }
                            else
                            {
                                return GetFormulaStringInner(ref ptgStack, 1, showAttributes);
                            }
                        case PtgAttrIf attrIf:
                            if (showAttributes)
                            {
                                return string.Format("[AttrIf:{0}]", attrIf.Offset) + GetFormulaStringInner(ref ptgStack, 1, showAttributes);
                            }
                            else
                            {
                                return GetFormulaStringInner(ref ptgStack, 1, showAttributes);
                            }
                        case PtgName ptgName:
                            if (string.IsNullOrEmpty(ptgName.nameValue))
                            {
                                return string.Format("PtgName(nameindex:{0})", ptgName.nameindex);
                            }
                            else
                            {
                                return ptgName.nameValue;
                            }
                            
                        default:
                            //Special case for operators
                            if (nextPtg is PtgParen)
                            {
                                return string.Format("({0})", GetFormulaStringInner(ref ptgStack, 1, showAttributes));
                            }
                            //If this is a simple operator like + or -
                            if (nextPtg.getLength() == 1)
                            {
                                //We pop the arguments off the stack in reverse order, hence the 'backwards' ordering
                                string formatString = "{2}{1}{0}";
                                return string.Format(formatString, GetFormulaStringInner(ref ptgStack, 1, showAttributes),
                                    nextPtg.getData(), GetFormulaStringInner(ref ptgStack, 1, showAttributes));
                            }
                            throw new NotImplementedException(string.Format("Ptg type {0} not implemented for ToString", nextPtg.Id));
                    }
                }
                else
                {
                    switch (nextPtg)
                    {
                        case PtgRef ptgRef: 
                            //Make sure to truncate column values that use any value > 255
                            return ExcelHelperClass.ConvertR1C1ToA1(string.Format("R{0}C{1}", ptgRef.rw + 1, (ptgRef.col + 1) & 0xFF));
                        case PtgArea ptgArea:
                            string firstCell = ExcelHelperClass.ConvertR1C1ToA1(string.Format("R{0}C{1}", ptgArea.rwFirst + 1, (ptgArea.colFirst + 1) & 0xFF));
                            string secondCell = ExcelHelperClass.ConvertR1C1ToA1(string.Format("R{0}C{1}", ptgArea.rwLast + 1, (ptgArea.colLast + 1) & 0xFF));
                            return string.Format("{0}:{1}", firstCell, secondCell);
                        case PtgStr ptgStr: return string.Format("\"{0}\"", ptgStr.getData());
                        default: return nextPtg.getData();
                    }
                    
                }
            }

            throw new NotImplementedException();
        }


        public static string GetFormulaString(Stack<AbstractPtg> ptgStack, bool showAttributes = false)
        {
            Stack<AbstractPtg> cloneStack = new Stack<AbstractPtg>(ptgStack.Reverse());
            string formulaString = "";
            formulaString  = GetFormulaStringInner(ref cloneStack, 1, showAttributes);

            if (cloneStack.Count > 0)
            {
                //Special case where SET.NAME is alias for a variable assignment
                if (cloneStack.Peek() is PtgAttrBaxcel)
                {
                    string regexStr = @"SET\.NAME\((.+?),(.*)\)";
                    Regex setNameRegex = new Regex(regexStr);
                    Match m = setNameRegex.Match(formulaString);
                    string varNameMatch = m.Groups[1].Value;
                    string valueNameMatch = m.Groups[2].Value;
                    string replacement = string.Format("{0}={1}", varNameMatch.Replace("\"", ""), valueNameMatch);
                    formulaString = formulaString.Replace(m.Groups[0].Value, replacement);
                    return formulaString;
                }
            }

            return formulaString;
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

                if (isByte19PtgHeader)
                {
                    bw.Write(Convert.ToByte(0x19));
                }
                else if (isByte18PtgHeader)
                {
                    bw.Write(Convert.ToByte(0x18));
                }
                else if ((int)ptgNumber >= 0x20)
                {
                    if (ptg.dataType == AbstractPtg.PtgDataType.VALUE) ptgNumber += 0x20;
                    if (ptg.dataType == AbstractPtg.PtgDataType.ARRAY) ptgNumber += 0x40;
                }

                //Allow explicitly setting the reserved bit in PtgNumbers for signature avoidance
                //NOTE: Excel appears to break if you do this, this functionality is kept for later experimentation
                if (ptg.fPtgNumberHighBit) ptgNumber += 0x80;

                bw.Write(Convert.ToByte(ptgNumber));
                bw.Write(GetByteDataFromPtg(ptg));
            }

            return bw.GetBytesWritten();
        }
    }
}
