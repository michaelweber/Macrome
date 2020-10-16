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

        private static Stack<AbstractPtg> UpdateNameRecords(Stack<AbstractPtg> ptgStack, List<Lbl> labelRecords)
        {
            List<AbstractPtg> modifiedStack = new List<AbstractPtg>();
            foreach (var ptg in ptgStack)
            {
                if (ptg is PtgName)
                {
                    int index = (ptg as PtgName).nameindex;
                    string matchingLabel = labelRecords[index - 1].Name.Value;
                    modifiedStack.Add(new PtgName(index, matchingLabel));
                }
                else if (ptg is PtgNameX)
                {
                    int index = (ptg as PtgNameX).nameindex;
                    ushort ixti = (ptg as PtgNameX).ixti;

                    string matchingLabel = labelRecords[index - 1].Name.Value;
                    modifiedStack.Add(new PtgNameX(ixti, index, matchingLabel));
                }
                else
                {
                    modifiedStack.Add(ptg);
                }
            }
            modifiedStack.Reverse();
            return new Stack<AbstractPtg>(modifiedStack);
        }

        private static Stack<AbstractPtg> UpdateSheetReferences(Stack<AbstractPtg> ptgStack, List<BoundSheet8> sheetRecords, ExternSheet externSheetRecord)
        {
            List<AbstractPtg> modifiedStack = new List<AbstractPtg>();
            foreach (var ptg in ptgStack)
            {
                if (ptg is PtgRef3d)
                {
                    PtgRef3d ref3d = (ptg as PtgRef3d);
                    int index = ref3d.ixti;
                    XTI relevantXti = externSheetRecord.rgXTI[index];
                    //Make sure this isn't a sheet or workbook level reference
                    if (relevantXti.itabFirst >= 0)
                    {
                        BoundSheet8 relevantSheet = sheetRecords[relevantXti.itabFirst];
                        string sheetName = relevantSheet.stName.Value;
                        modifiedStack.Add(new PtgRef3d(ref3d.rw, ref3d.col, ref3d.ixti, ref3d.rwRelative, ref3d.colRelative, sheetName));
                    }
                    else
                    {
                        modifiedStack.Add(ptg);
                    }
                }
                else if (ptg is PtgNameX)
                {
                    PtgNameX ptgNameX = (ptg as PtgNameX);
                    int index = ptgNameX.ixti - 1;
                    XTI relevantXti = externSheetRecord.rgXTI[index];

                    if (relevantXti.itabFirst >= 0 && relevantXti.itabFirst < sheetRecords.Count)
                    {
                        BoundSheet8 relevantSheet = sheetRecords[relevantXti.itabFirst];
                        string sheetName = relevantSheet.stName.Value;
                        modifiedStack.Add(new PtgNameX(ptgNameX.ixti, ptgNameX.nameindex, ptgNameX.nameValue, sheetName));
                    }
                    else
                    {
                        modifiedStack.Add(ptg);
                    }
                }
                else
                {
                    modifiedStack.Add(ptg);
                }
            }
            modifiedStack.Reverse();
            return new Stack<AbstractPtg>(modifiedStack);
        }

        public static List<BiffRecord> UpdateGlobalsStreamReferences(List<BiffRecord> records)
        {
            List<Lbl> labelRecords = records.Where(r => r.Id == RecordType.Lbl).Select(r => r.AsRecordType<Lbl>()).ToList();
            List<BoundSheet8> sheetRecords = records.Where(r => r.Id == RecordType.BoundSheet8).Select(r => r.AsRecordType<BoundSheet8>()).ToList();
            //Incorrect way to do this, but works for simpler cases - just pull the first ExternSheet record in the XLS file.
            BiffRecord firstExternSheet = records.FirstOrDefault(r => r.Id == RecordType.ExternSheet);

            List<BiffRecord> updatedRecords = new List<BiffRecord>();
            foreach (var record in records)
            {
                switch (record)
                {
                    case Formula formulaRecord:
                        Stack<AbstractPtg> modifiedFormulaStack = UpdateNameRecords(formulaRecord.ptgStack, labelRecords);
                        if (firstExternSheet != null)
                        {
                            modifiedFormulaStack = UpdateSheetReferences(modifiedFormulaStack, sheetRecords,
                                firstExternSheet.AsRecordType<ExternSheet>());
                        }
                        formulaRecord.SetCellParsedFormula(new CellParsedFormula(modifiedFormulaStack));
                        updatedRecords.Add(formulaRecord);
                        continue;
                    case Lbl lblRecord:
                        Stack<AbstractPtg> modifiedLabelStack = UpdateNameRecords(lblRecord.rgce, labelRecords);
                        if (firstExternSheet != null)
                        {
                            modifiedLabelStack = UpdateSheetReferences(modifiedLabelStack, sheetRecords,
                                firstExternSheet.AsRecordType<ExternSheet>());
                        }
                        lblRecord.SetRgce(modifiedLabelStack);
                        updatedRecords.Add(lblRecord);
                        continue;
                    default:
                        updatedRecords.Add(record);
                        continue;
                }

            }
            return updatedRecords;
        }

        private static byte[] GetByteDataFromPtg(AbstractPtg ptg, bool isAttributePtg = false)
        {
            MemoryStream ms = new MemoryStream();
            BinaryWriter bw = new BinaryWriter(ms);

            if (isAttributePtg)
            {
                switch (ptg)
                {
                    case PtgAttrBaxcel ptgAttrBaxcel:
                        bw.Write(ptgAttrBaxcel.DataBytes);
                        break;
                    case PtgAttrIf ptgAttrIf:
                        bw.Write(ptgAttrIf.PtgOffset);
                        break;
                    case PtgAttrGoto ptgAttrGoto:
                        bw.Write(ptgAttrGoto.PtgOffset);
                        break;
                    case PtgAttrSpace ptgAttrSpace:
                        bw.Write(ptgAttrSpace.PtgAttrSpaceType);
                        break;
                    case PtgAttrSemi ptgAttrSemi:
                        bw.Write(ptgAttrSemi.Unused);
                        break;
                    case PtgAttrSum ptgAttrSum:
                        bw.Write(ptgAttrSum.Unused);
                        break;
                    default:
                        throw new NotImplementedException(string.Format("No byte conversion implemented for {0}", (Ptg0x19Sub)ptg.Id));
                }

                return bw.GetBytesWritten();
            }

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
                case PtgRange ptgRange:
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
                
                case PtgName ptgName:
                    bw.Write(ptgName.nameindex);
                    break;
                case PtgNameX ptgNameX:
                    bw.Write(Convert.ToUInt16(ptgNameX.ixti));
                    bw.Write(Convert.ToUInt32(ptgNameX.nameindex));
                    break;
                case PtgMemArea ptgMemArea:
                    bw.Write(ptgMemArea.Unused);
                    bw.Write(ptgMemArea.cce);
                    break;
                case PtgArea ptgArea:
                    bw.Write(ptgArea.rwFirst);
                    bw.Write(ptgArea.rwLast);
                    bw.Write(ptgArea.colFirst);
                    bw.Write(ptgArea.colLast);
                    break;
                case PtgArea3d ptgArea3d:
                    bw.Write(ptgArea3d.ixti);
                    bw.Write(ptgArea3d.rwFirst);
                    bw.Write(ptgArea3d.rwLast);
                    bw.Write(ptgArea3d.colFirst);
                    bw.Write(ptgArea3d.colLast);
                    break;

                case PtgErr ptgErr:
                    bw.Write(ptgErr.Err);
                    break;
                case PtgAreaErr3d ptgAreaErr3d:
                    bw.Write(ptgAreaErr3d.ixti);
                    bw.Write(ptgAreaErr3d.unused1);
                    bw.Write(ptgAreaErr3d.unused2);
                    bw.Write(ptgAreaErr3d.unused3);
                    bw.Write(ptgAreaErr3d.unused4);
                    break;
                case PtgMemFunc ptgMemFunc:
                    bw.Write(ptgMemFunc.cce);
                    bw.Write(GetBytes(ptgMemFunc.ptgStack));
                    break;
                //Start 0x19 ## Section

                case PtgAttrChoose ptgAttrChoose:
                
                case PtgNotDocumented ptgNotDocumented:
                case PtgMemErr ptgMemErr:
                
                case PtgAreaN ptgAreaN:
                
                
                case PtgRefErr ptgRefErr:
                case PtgRefErr3d ptgRefErr3d:
                case PtgAreaErr ptgAreaErr:
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
                args.Push(GetFormulaStringInner(ref ptgStack, showAttributes));
            }

            return args.ToList();
        }

        private static string GetFormulaStringInner(ref Stack<AbstractPtg> ptgStack, bool showAttributes = false)
        {
            if (ptgStack.Count == 0)
            {
                Console.WriteLine("ERROR! Stack is Empty");
                return "EMPTYSTACKERROR";
            }

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
                            formulaVariableString = GetFormulaStringInner(ref ptgStack, showAttributes);
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
                            return string.Format("[AttrBaxcel:{0}]", attrBaxcel.DataBytes) + GetFormulaStringInner(ref ptgStack, showAttributes);
                        }
                        else
                        {
                            return GetFormulaStringInner(ref ptgStack, showAttributes);
                        }
                    case PtgAttrSpace attrSpace:
                        if (showAttributes)
                        {
                            return string.Format("[AttrSpace:{0}",
                                attrSpace.PtgAttrSpaceType + GetFormulaStringInner(ref ptgStack, showAttributes));
                        }
                        else
                        {
                            return GetFormulaStringInner(ref ptgStack, showAttributes);
                        }
                    case PtgAttrGoto attrGoto:
                        if (showAttributes)
                        {
                            return string.Format("[AttrGoto:{0}]", attrGoto.PtgOffset) + GetFormulaStringInner(ref ptgStack, showAttributes);
                        }
                        else
                        {
                            return GetFormulaStringInner(ref ptgStack, showAttributes);
                        }
                    case PtgAttrIf attrIf:
                        if (showAttributes)
                        {
                            return string.Format("[AttrIf:{0}]", attrIf.PtgOffset) + GetFormulaStringInner(ref ptgStack, showAttributes);
                        }
                        else
                        {
                            return GetFormulaStringInner(ref ptgStack, showAttributes);
                        }
                    case PtgAttrSemi attrSemi:
                        if (showAttributes)
                        {
                            return string.Format("[AttrSemi:{0}]", attrSemi.Unused) +
                                   GetFormulaStringInner(ref ptgStack, showAttributes);
                        }
                        else
                        {
                            return GetFormulaStringInner(ref ptgStack, showAttributes);
                        }
                    case PtgAttrSum attrSum:
                        if (showAttributes)
                        {
                            return string.Format("[AttrSum:{0}]", attrSum.Unused) +
                                   GetFormulaStringInner(ref ptgStack, showAttributes);
                        }
                        else
                        {
                            return GetFormulaStringInner(ref ptgStack, showAttributes);
                        }
                    case PtgName ptgName:
                        return ptgName.ToString();
                    case PtgNameX ptgNameX:
                        return ptgNameX.ToString();
                    case PtgBool ptgBool:
                        return ptgBool.getData();
                    default:
                        //Special case for operators
                        if (nextPtg is PtgParen)
                        {
                            return string.Format("({0})", GetFormulaStringInner(ref ptgStack, showAttributes));
                        }
                        else if (nextPtg is PtgUminus || nextPtg is PtgUplus || nextPtg is PtgPercent)
                        {
                            return string.Format("{1}{0}", GetFormulaStringInner(ref ptgStack, showAttributes),
                                nextPtg.getData());
                        }
                        //If this is a simple operator like + or -
                        else if (nextPtg.getLength() == 1)
                        {
                            //We pop the arguments off the stack in reverse order, hence the 'backwards' ordering
                            string formatString = "{2}{1}{0}";
                            return string.Format(formatString, GetFormulaStringInner(ref ptgStack, showAttributes),
                                nextPtg.getData(), GetFormulaStringInner(ref ptgStack, showAttributes));
                        }
                        throw new NotImplementedException(string.Format("Ptg type {0} not implemented for ToString", nextPtg.Id));
                }
            }
            else
            {
                switch (nextPtg)
                {
                    case PtgRef ptgRef:
                        return ptgRef.ToString();
                    case PtgArea ptgArea:
                        return ptgArea.ToString();
                    case PtgArea3d ptgArea3d:
                        return ptgArea3d.ToString();
                    case PtgMemArea ptgMemArea:
                        return "MEMAREA";
                    case PtgStr ptgStr: return string.Format("\"{0}\"", ptgStr.getData());
                    default: return nextPtg.getData();
                }
                
            }
        }


        public static string GetFormulaString(Stack<AbstractPtg> ptgStack, bool showAttributes = false)
        {
            if (ptgStack.Count == 0)
            {
                return "!Null Ptg Stack! (Manually Edited Ptg Record)";
            }

            Stack<AbstractPtg> cloneStack = new Stack<AbstractPtg>(ptgStack.Reverse());
            string formulaString = "";
            formulaString  = GetFormulaStringInner(ref cloneStack, showAttributes);

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
                bw.Write(GetByteDataFromPtg(ptg, isByte19PtgHeader));
            }

            return bw.GetBytesWritten();
        }
    }
}
