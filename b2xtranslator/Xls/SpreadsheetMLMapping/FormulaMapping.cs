

using System;
using System.Collections.Generic;
using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Ptg;
using b2xtranslator.Tools;

namespace b2xtranslator.SpreadsheetMLMapping
{
    public class FormulaInfixMapping
    {

        /// <summary>
        /// This static method is used to convert the ptgStack to the infixnotation
        /// Normal parsing without row and column changes 
        /// </summary>
        /// <param name="stack"></param>
        /// <param name="xlsContext"></param>
        /// <returns></returns>
        public static string mapFormula(Stack<AbstractPtg> stack, ExcelContext xlsContext)
        {
            return FormulaInfixMapping.mapFormula(stack, xlsContext, 0, 0);
        }

        /// <summary>
        /// This static method is used to convert the ptgStack to the infixnotation
        /// This method changes every Ptg***N** PtgRecord from a shared formula
        /// </summary>
        /// <param name="stack"></param>
        /// <param name="xlsContext"></param>
        /// <param name="rw">Row from the shared formula</param>
        /// <param name="col">Column from  the shared formula</param>
        /// <returns></returns>
        public static string mapFormula(Stack<AbstractPtg> stack, ExcelContext xlsContext, int rw, int col)
        {
            var resultStack = new Stack<string>();
            try
            {
                var opStack = new Stack<AbstractPtg>(stack);
                if (opStack.Count == 0)
                {
                    throw new Exception();
                }

                while (opStack.Count != 0)
                {
                    var ptg = opStack.Pop();

                    if (ptg is PtgInt || ptg is PtgNum || ptg is PtgBool || ptg is PtgMissArg || ptg is PtgErr)
                    {
                        resultStack.Push(ptg.getData());
                    }
                    else if (ptg is PtgRef)
                    {
                        var ptgref = (PtgRef)ptg;

                        int realCol = ptgref.col;
                        int realRw = ptgref.rw + 1;

                        if (ptgref.colRelative)
                        {
                            realCol += col;
                        }

                        if (ptgref.rwRelative)
                        {
                            realRw += rw;
                        }

                        resultStack.Push(ExcelHelperClass.intToABCString(realCol, (realRw).ToString(), ptgref.colRelative, ptgref.rwRelative));
                    }
                    else if (ptg is PtgRefN)
                    {
                        var ptgrefn = (PtgRefN)ptg;
                        int realCol = (int)ptgrefn.col;
                        int realRw;

                        if (ptgrefn.colRelative)
                        {
                            realCol += col;
                        }
                        if (realCol >= 0xFF)
                        {
                            realCol -= 0x0100;
                        }

                        if (ptgrefn.rwRelative)
                        {
                            realRw = (short)ptgrefn.rw + 1;
                            realRw += rw;
                        }
                        else
                        {
                            realRw = ptgrefn.rw + 1;
                        }

                        resultStack.Push(ExcelHelperClass.intToABCString(realCol, (realRw).ToString(), ptgrefn.colRelative, ptgrefn.rwRelative));
                    }
                    else if (ptg is PtgUplus || ptg is PtgUminus)
                    {
                        string buffer = "";
                        if (ptg.PopSize() > resultStack.Count)
                        {
                            throw new ExtractorException(ExtractorException.PARSEDFORMULAEXCEPTION);
                        }
                        buffer = ptg.getData() + resultStack.Pop();

                        resultStack.Push(buffer);
                    }
                    else if (ptg is PtgParen)
                    {
                        string buffer = "";
                        if (ptg.PopSize() > resultStack.Count)
                        {
                            throw new ExtractorException(ExtractorException.PARSEDFORMULAEXCEPTION);
                        }
                        buffer = "(" + resultStack.Pop() + ")";

                        resultStack.Push(buffer);
                    }
                    else if (ptg is PtgPercent)
                    {
                        string buffer = "";
                        if (ptg.PopSize() > resultStack.Count)
                        {
                            throw new ExtractorException(ExtractorException.PARSEDFORMULAEXCEPTION);
                        }
                        buffer = resultStack.Pop() + ptg.getData();

                        resultStack.Push(buffer);
                    }
                    else if (ptg is PtgAdd || ptg is PtgDiv || ptg is PtgMul ||
                                ptg is PtgSub || ptg is PtgPower || ptg is PtgGt ||
                                ptg is PtgGe || ptg is PtgLt || ptg is PtgLe ||
                                ptg is PtgEq || ptg is PtgNe || ptg is PtgConcat ||
                                ptg is PtgUnion || ptg is PtgIsect
                        )
                    {
                        string buffer = "";
                        if (ptg.PopSize() > resultStack.Count)
                        {
                            throw new ExtractorException(ExtractorException.PARSEDFORMULAEXCEPTION);
                        }
                        buffer = ptg.getData() + resultStack.Pop();
                        buffer = resultStack.Pop() + buffer;
                        resultStack.Push(buffer);
                    }
                    else if (ptg is PtgStr)
                    {
                        resultStack.Push("\"" + ptg.getData() + "\"");
                    }
                    else if (ptg is PtgArea)
                    {
                        string buffer = "";
                        var ptgarea = (PtgArea)ptg;
                        buffer = ExcelHelperClass.intToABCString((int)ptgarea.colFirst, (ptgarea.rwFirst + 1).ToString(), ptgarea.colFirstRelative, ptgarea.rwFirstRelative);
                        buffer = buffer + ":" + ExcelHelperClass.intToABCString((int)ptgarea.colLast, (ptgarea.rwLast + 1).ToString(), ptgarea.colLastRelative, ptgarea.rwLastRelative);


                        resultStack.Push(buffer);
                    }
                    else if (ptg is PtgAreaN)
                    {
                        string buffer = "";
                        var ptgarean = (PtgAreaN)ptg;
                        int realRwFirst;
                        int realRwLast;
                        int realColFirst = (int)ptgarean.colFirst;
                        int realColLast = (int)ptgarean.colLast;

                        if (ptgarean.colFirstRelative)
                        {
                            realColFirst += col;
                        }

                        if (realColFirst >= 0xFF)
                        {
                            realColFirst -= 0x0100;
                        }

                        if (ptgarean.colLastRelative)
                        {
                            realColLast += col;
                        }
                        if (realColLast >= 0xFF)
                        {
                            realColLast -= 0x0100;
                        }

                        if (ptgarean.rwFirstRelative)
                        {
                            realRwFirst = (short)ptgarean.rwFirst + 1;
                            realRwFirst += rw;
                        }
                        else
                        {
                            realRwFirst = ptgarean.rwFirst + 1;
                        }
                        if (ptgarean.rwLastRelative)
                        {
                            realRwLast = (short)ptgarean.rwLast + 1;
                            realRwLast += rw;
                        }
                        else
                        {
                            realRwLast = ptgarean.rwLast + 1;
                        }

                        buffer = ExcelHelperClass.intToABCString(realColFirst, (realRwFirst).ToString(), ptgarean.colFirstRelative, ptgarean.rwFirstRelative);
                        buffer = buffer + ":" + ExcelHelperClass.intToABCString(realColLast, (realRwLast).ToString(), ptgarean.colLastRelative, ptgarean.rwLastRelative);


                        resultStack.Push(buffer);
                    }
                    else if (ptg is PtgAttrSum)
                    {
                        string buffer = "";
                        var ptgref = (PtgAttrSum)ptg;
                        buffer = ptg.getData() + "(" + resultStack.Pop() + ")";
                        resultStack.Push(buffer);
                    }
                    else if (ptg is PtgAttrIf || ptg is PtgAttrGoto || ptg is PtgAttrSemi
                    || ptg is PtgAttrChoose || ptg is PtgAttrSpace)
                    {

                    }
                    else if (ptg is PtgExp)
                    {
                        var ptgexp = (PtgExp)ptg;
                        var newptgstack = ((WorkSheetData)xlsContext.CurrentSheet).getArrayData(ptgexp.rw, ptgexp.col);
                        resultStack.Push(FormulaInfixMapping.mapFormula(newptgstack, xlsContext));
                        ((WorkSheetData)xlsContext.CurrentSheet).setFormulaUsesArray(ptgexp.rw, ptgexp.col);

                    }
                    else if (ptg is PtgRef3d)
                    {
                        try
                        {
                            var ptgr3d = (PtgRef3d)ptg;
                            string refstring = ExcelHelperClass.EscapeFormulaString(xlsContext.XlsDoc.WorkBookData.getIXTIString(ptgr3d.ixti));
                            string cellref = ExcelHelperClass.intToABCString((int)ptgr3d.col, (ptgr3d.rw + 1).ToString(), ptgr3d.colRelative, ptgr3d.rwRelative);

                            resultStack.Push("'" + refstring + "'" + "!" + cellref);
                        }
                        catch (Exception)
                        {
                            resultStack.Push("#REF!");
                        }
                    }
                    else if (ptg is PtgArea3d)
                    {
                        try
                        {
                            var ptga3d = (PtgArea3d)ptg;
                            string refstring = ExcelHelperClass.EscapeFormulaString(xlsContext.XlsDoc.WorkBookData.getIXTIString(ptga3d.ixti));
                            string buffer = "";
                            buffer = ExcelHelperClass.intToABCString((int)ptga3d.colFirst, (ptga3d.rwFirst + 1).ToString(), ptga3d.colFirstRelative, ptga3d.rwFirstRelative);
                            buffer = buffer + ":" + ExcelHelperClass.intToABCString((int)ptga3d.colLast, (ptga3d.rwLast + 1).ToString(), ptga3d.colLastRelative, ptga3d.rwLastRelative);

                            resultStack.Push("'" + refstring + "'!" + buffer);
                        }
                        catch (Exception)
                        {
                            resultStack.Push("#REF!");
                        }

                    }
                    else if (ptg is PtgNameX)
                    {
                        var ptgnx = (PtgNameX)ptg;
                        string opstring = xlsContext.XlsDoc.WorkBookData.getExternNameByRef(ptgnx.ixti, ptgnx.nameindex);
                        resultStack.Push(opstring);
                    }
                    else if (ptg is PtgName)
                    {
                        var ptgn = (PtgName)ptg;
                        string opstring = xlsContext.XlsDoc.WorkBookData.getDefinedNameByRef(ptgn.nameindex);
                        resultStack.Push(opstring);
                    }
                    else if (ptg is PtgRefErr)
                    {
                        var ptgreferr = (PtgRefErr)ptg;
                        resultStack.Push(ptgreferr.getData());
                    }
                    else if (ptg is PtgRefErr3d)
                    {
                        try
                        {
                            var ptgreferr3d = (PtgRefErr3d)ptg;

                            string refstring = ExcelHelperClass.EscapeFormulaString(xlsContext.XlsDoc.WorkBookData.getIXTIString(ptgreferr3d.ixti));
                            resultStack.Push("'" + refstring + "'" + "!" + ptgreferr3d.getData());
                        }
                        catch (Exception)
                        {
                            resultStack.Push("#REF!");
                        }
                    }
                    else if (ptg is PtgAreaErr3d)
                    {
                        try
                        {
                            var ptgareaerr3d = (PtgAreaErr3d)ptg;
                            string refstring = ExcelHelperClass.EscapeFormulaString(xlsContext.XlsDoc.WorkBookData.getIXTIString(ptgareaerr3d.ixti));
                            resultStack.Push("'" + refstring + "'" + "!" + ptgareaerr3d.getData());
                        }
                        catch (Exception)
                        {
                            resultStack.Push("#REF!");
                        }
                    }
                    else if (ptg is PtgFunc)
                    {
                        var ptgf = (PtgFunc)ptg;
                        var value = (FtabValues)ptgf.tab;
                        string buffer = value.ToString();
                        buffer.Replace("_", ".");

                        // no param 
                        if (value == FtabValues.NA || value == FtabValues.PI ||
                            value == FtabValues.TRUE || value == FtabValues.FALSE
                            || value == FtabValues.RAND || value == FtabValues.NOW
                            || value == FtabValues.TODAY
                            )
                        {
                            buffer += "()";
                        }
                        // One param 
                        else if (value == FtabValues.ISNA || value == FtabValues.ISERROR ||
                                    value == FtabValues.SIN || value == FtabValues.COS ||
                            value == FtabValues.TAN || value == FtabValues.ATAN ||
                            value == FtabValues.SQRT || value == FtabValues.EXP ||
                            value == FtabValues.LN || value == FtabValues.LOG10 ||
                            value == FtabValues.ABS || value == FtabValues.INT ||
                            value == FtabValues.SIGN || value == FtabValues.LEN ||
                            value == FtabValues.VALUE || value == FtabValues.NOT
                            || value == FtabValues.DAY || value == FtabValues.MONTH
                            || value == FtabValues.YEAR || value == FtabValues.HOUR
                            || value == FtabValues.MINUTE || value == FtabValues.SECOND
                            || value == FtabValues.AREAS || value == FtabValues.ROWS
                            || value == FtabValues.COLUMNS || value == FtabValues.TRANSPOSE
                            || value == FtabValues.TYPE || value == FtabValues.ASIN
                            || value == FtabValues.ACOS || value == FtabValues.ISREF
                            || value == FtabValues.CHAR || value == FtabValues.LOWER
                            || value == FtabValues.UPPER || value == FtabValues.PROPER
                            || value == FtabValues.TRIM || value == FtabValues.CODE
                            || value == FtabValues.ISERR || value == FtabValues.ISTEXT
                            || value == FtabValues.ISNUMBER || value == FtabValues.ISBLANK
                            || value == FtabValues.T || value == FtabValues.N
                            || value == FtabValues.DATEVALUE || value == FtabValues.TIMEVALUE
                            || value == FtabValues.CLEAN || value == FtabValues.MDETERM
                            || value == FtabValues.MINVERSE || value == FtabValues.FACT
                            || value == FtabValues.ISNONTEXT || value == FtabValues.ISLOGICAL
                            || value == FtabValues.LENB || value == FtabValues.DBCS
                            || value == FtabValues.SINH || value == FtabValues.COSH
                            || value == FtabValues.TANH || value == FtabValues.ASINH
                            || value == FtabValues.ACOSH || value == FtabValues.ATANH
                            || value == FtabValues.INFO || value == FtabValues.ERROR_TYPE
                            || value == FtabValues.GAMMALN || value == FtabValues.EVEN
                            || value == FtabValues.FISHER || value == FtabValues.FISHERINV
                            || value == FtabValues.NORMSDIST || value == FtabValues.NORMSINV
                            || value == FtabValues.ODD || value == FtabValues.RADIANS
                            || value == FtabValues.DEGREES || value == FtabValues.COUNTBLANK
                            || value == FtabValues.DATESTRING || value == FtabValues.PHONETIC
                            )
                        {
                            buffer += "(" + resultStack.Pop() + ")";
                        }
                        // two params 
                        else if (value == FtabValues.ROUND || value == FtabValues.REPT ||
                                value == FtabValues.MOD || value == FtabValues.TEXT
                                || value == FtabValues.ATAN2 || value == FtabValues.EXACT
                                || value == FtabValues.MMULT || value == FtabValues.ROUNDUP
                                || value == FtabValues.ROUNDDOWN || value == FtabValues.FREQUENCY
                                || value == FtabValues.CHIDIST || value == FtabValues.CHIINV
                                || value == FtabValues.COMBIN || value == FtabValues.FLOOR
                                || value == FtabValues.CEILING || value == FtabValues.PERMUT
                                || value == FtabValues.SUMXMY2 || value == FtabValues.SUMX2MY2
                                || value == FtabValues.SUMX2PY2 || value == FtabValues.CHITEST
                                || value == FtabValues.CORREL || value == FtabValues.COVAR
                                || value == FtabValues.FTEST || value == FtabValues.INTERCEPT
                                || value == FtabValues.PEARSON || value == FtabValues.RSQ
                                || value == FtabValues.STEYX || value == FtabValues.SLOPE
                                || value == FtabValues.LARGE || value == FtabValues.SMALL
                                || value == FtabValues.QUARTILE || value == FtabValues.PERCENTILE
                                || value == FtabValues.TRIMMEAN || value == FtabValues.TINV
                                || value == FtabValues.POWER || value == FtabValues.COUNTIF
                            || value == FtabValues.NUMBERSTRING


                            )
                        {
                            buffer += "(";
                            string buffer2 = resultStack.Pop();
                            buffer2 = resultStack.Pop() + "," + buffer2;

                            buffer += buffer2 + ")";
                        }
                        // Three params 
                        else if (value == FtabValues.MID || value == FtabValues.DCOUNT ||
                                    value == FtabValues.DSUM || value == FtabValues.DAVERAGE ||
                                    value == FtabValues.DMIN || value == FtabValues.DMAX ||
                                    value == FtabValues.DSTDEV || value == FtabValues.DVAR
                                || value == FtabValues.MIRR || value == FtabValues.DATE
                                || value == FtabValues.TIME || value == FtabValues.SLN
                                || value == FtabValues.DPRODUCT || value == FtabValues.DSTDEVP
                                || value == FtabValues.DVARP || value == FtabValues.DCOUNTA
                                || value == FtabValues.MIDB || value == FtabValues.DGET
                                || value == FtabValues.CONFIDENCE || value == FtabValues.CRITBINOM
                                || value == FtabValues.EXPONDIST || value == FtabValues.FDIST

                                || value == FtabValues.FINV || value == FtabValues.GAMMAINV
                                || value == FtabValues.LOGNORMDIST || value == FtabValues.LOGINV
                                || value == FtabValues.NEGBINOMDIST || value == FtabValues.NORMINV
                                || value == FtabValues.STANDARDIZE || value == FtabValues.POISSON
                                || value == FtabValues.TDIST || value == FtabValues.FORECAST
                            )
                        {
                            buffer += "(";
                            string buffer2 = resultStack.Pop();
                            buffer2 = resultStack.Pop() + "," + buffer2;
                            buffer2 = resultStack.Pop() + "," + buffer2;
                            buffer += buffer2 + ")";
                        }
                        // four params 
                        else if (value == FtabValues.REPLACE || value == FtabValues.SYD
                                || value == FtabValues.REPLACEB || value == FtabValues.BINOMDIST
                                || value == FtabValues.GAMMADIST || value == FtabValues.HYPGEOMDIST
                                || value == FtabValues.NORMDIST || value == FtabValues.WEIBULL
                                || value == FtabValues.TTEST || value == FtabValues.ISPMT

                            )
                        {
                            buffer += "(";
                            string buffer2 = resultStack.Pop();
                            buffer2 = resultStack.Pop() + "," + buffer2;
                            buffer2 = resultStack.Pop() + "," + buffer2;
                            buffer2 = resultStack.Pop() + "," + buffer2;
                            buffer += buffer2 + ")";
                        }
                        if ((int)value != 0xff)
                        {

                        }
                        resultStack.Push(buffer);

                    }
                    else if (ptg is PtgFuncVar)
                    {
                        var ptgfv = (PtgFuncVar)ptg;
                        // is Ftab or Cetab 
                        if (!ptgfv.fCelFunc)
                        {
                            var value = (FtabValues)ptgfv.tab;
                            string buffer = value.ToString();
                            buffer.Replace("_", ".");
                            // 1 to x parameter
                            if (value == FtabValues.COUNT || value == FtabValues.IF ||
                                value == FtabValues.ISNA || value == FtabValues.ISERROR ||
                                value == FtabValues.AVERAGE || value == FtabValues.MAX ||
                                value == FtabValues.MIN || value == FtabValues.SUM ||
                                value == FtabValues.ROW || value == FtabValues.COLUMN ||
                                value == FtabValues.NPV || value == FtabValues.STDEV ||
                                value == FtabValues.DOLLAR || value == FtabValues.FIXED ||
                                value == FtabValues.SIN || value == FtabValues.COS ||
                                value == FtabValues.LOOKUP || value == FtabValues.INDEX ||
                                value == FtabValues.AND || value == FtabValues.OR ||
                                value == FtabValues.VAR || value == FtabValues.LINEST ||
                                value == FtabValues.TREND || value == FtabValues.LOGEST ||
                                value == FtabValues.GROWTH || value == FtabValues.PV ||
                                value == FtabValues.FV || value == FtabValues.NPER ||
                                value == FtabValues.PMT || value == FtabValues.RATE ||
                                value == FtabValues.IRR || value == FtabValues.MATCH ||
                                value == FtabValues.WEEKDAY || value == FtabValues.OFFSET
                                || value == FtabValues.ARGUMENT || value == FtabValues.SEARCH
                                || value == FtabValues.CHOOSE || value == FtabValues.HLOOKUP
                                || value == FtabValues.VLOOKUP || value == FtabValues.LOG
                                || value == FtabValues.LEFT || value == FtabValues.RIGHT
                                || value == FtabValues.SUBSTITUTE || value == FtabValues.FIND
                                || value == FtabValues.CELL || value == FtabValues.DDB
                                || value == FtabValues.INDIRECT || value == FtabValues.IPMT
                                || value == FtabValues.PPMT || value == FtabValues.COUNTA
                                || value == FtabValues.PRODUCT || value == FtabValues.STDEVP
                                || value == FtabValues.VARP || value == FtabValues.TRUNC
                                || value == FtabValues.USDOLLAR || value == FtabValues.FINDB
                                || value == FtabValues.SEARCHB || value == FtabValues.LEFTB
                                || value == FtabValues.RIGHTB || value == FtabValues.RANK
                                || value == FtabValues.ADDRESS || value == FtabValues.DAYS360
                                || value == FtabValues.VDB || value == FtabValues.MEDIAN
                                || value == FtabValues.PRODUCT || value == FtabValues.DB
                                || value == FtabValues.AVEDEV || value == FtabValues.BETADIST
                                || value == FtabValues.BETAINV || value == FtabValues.PROB
                                || value == FtabValues.DEVSQ || value == FtabValues.GEOMEAN
                                || value == FtabValues.HARMEAN || value == FtabValues.SUMSQ
                                || value == FtabValues.KURT || value == FtabValues.SKEW
                                || value == FtabValues.ZTEST || value == FtabValues.PERCENTRANK
                                || value == FtabValues.MODE || value == FtabValues.CONCATENATE
                                || value == FtabValues.SUBTOTAL || value == FtabValues.SUMIF
                                || value == FtabValues.ROMAN || value == FtabValues.GETPIVOTDATA
                                || value == FtabValues.HYPERLINK || value == FtabValues.AVERAGEA
                                || value == FtabValues.MAXA || value == FtabValues.MINA
                                || value == FtabValues.STDEVPA || value == FtabValues.VARPA
                                || value == FtabValues.STDEVA || value == FtabValues.VARA
                                )
                            {
                                buffer += "(";
                                string buffer2 = resultStack.Pop();
                                for (int i = 1; i < ptgfv.cparams; i++)
                                {
                                    buffer2 = resultStack.Pop() + "," + buffer2;
                                }
                                buffer += buffer2 + ")";
                                resultStack.Push(buffer);
                            }
                            else if ((int)value == 0xFF)
                            {
                                buffer = "(";
                                string buffer2 = resultStack.Pop();
                                for (int i = 1; i < ptgfv.cparams - 1; i++)
                                {
                                    buffer2 = resultStack.Pop() + "," + buffer2;
                                }
                                buffer += buffer2 + ")";
                                // take the additional Operator from the Operandstack 
                                buffer = resultStack.Pop() + buffer;
                                resultStack.Push(buffer);
                            }

                        }
                    }
                    else if (ptg is PtgMemFunc)
                    {
                        string value;
                        value = FormulaInfixMapping.mapFormula(((PtgMemFunc)ptg).ptgStack, xlsContext);
                        resultStack.Push(value);
                    }
                }
            }
            catch (Exception ex)
            {
                TraceLogger.DebugInternal(ex.Message.ToString());
                resultStack.Push("");
            }
            if (resultStack.Count == 0)
                resultStack.Push("");

            return resultStack.Pop();
        }

        /// <summary>
        /// returns the error string for the formula file 
        /// </summary>
        /// <param name="errorcode"></param>
        /// <returns></returns>
        public static string getErrorStringfromCode(int errorcode)
        {

            switch (errorcode)
            {
                case 0x00: return "#NULL!";
                case 0x07: return "#DIV/0!";
                case 0x0F: return "#VALUE!";
                case 0x17: return "#REF!";
                case 0x1D: return "#NAME?";
                case 0x24: return "#NUM!";
                case 0x2A: return "#N/A";
                default: return "#NULL!";
            }
        }
    }
}
