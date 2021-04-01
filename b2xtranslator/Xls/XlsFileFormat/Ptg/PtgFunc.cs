
using System;
using System.Diagnostics;
using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Ptg
{
    public class PtgFunc : AbstractPtg
    {
        public static int GetParamCountFromFtabValue(FtabValues value)
        {
            switch (value)
            {
                case FtabValues.NEXT:
                case FtabValues.ACTIVE_CELL:
                case FtabValues.NA:
                case FtabValues.PI:
                case FtabValues.TRUE:
                case FtabValues.FALSE:
                case FtabValues.RAND:
                case FtabValues.NOW:
                case FtabValues.STEP:
                case FtabValues.CALLER:
                case FtabValues.SELECTION:
                case FtabValues.BREAK:
                case FtabValues.TODAY:
                case FtabValues.ELSE:
                case FtabValues.END_IF:
                case FtabValues.LAST_ERROR:
                case FtabValues.GROUP:
                    return 0;
                case FtabValues.ISERROR:
                case FtabValues.WHILE:
                case FtabValues.GOTO:
                case FtabValues.LEN:
                case FtabValues.ISNA:
                case FtabValues.ISNUMBER:
                case FtabValues.GET_WORKSPACE:
                case FtabValues.CHAR:
                case FtabValues.SIN:
                case FtabValues.COS:
                case FtabValues.TAN:
                case FtabValues.ATAN:
                case FtabValues.SQRT:
                case FtabValues.LN:
                case FtabValues.LOG10:
                case FtabValues.ABS:
                case FtabValues.INT:
                case FtabValues.SIGN:
                case FtabValues.VALUE:
                case FtabValues.NOT:
                case FtabValues.DAY:
                case FtabValues.MONTH:
                case FtabValues.YEAR:
                case FtabValues.HOUR:
                case FtabValues.MINUTE:
                case FtabValues.SECOND:
                case FtabValues.AREAS:
                case FtabValues.ROWS:
                case FtabValues.COLUMNS:
                case FtabValues.TRANSPOSE:
                case FtabValues.TYPE:
                case FtabValues.DEREF:
                case FtabValues.ASIN:
                case FtabValues.ACOS:
                case FtabValues.ISREF:
                case FtabValues.GET_FORMULA:
                case FtabValues.LOWER:
                case FtabValues.UPPER:
                case FtabValues.PROPER:
                case FtabValues.TRIM:
                case FtabValues.CODE:
                case FtabValues.ISERR:
                case FtabValues.ISTEXT:
                case FtabValues.ISBLANK:
                case FtabValues.T:
                case FtabValues.N:
                case FtabValues.FCLOSE:
                case FtabValues.FSIZE:
                case FtabValues.FREADLN:
                case FtabValues.DATEVALUE:
                case FtabValues.TIMEVALUE:
                case FtabValues.DIALOG_BOX:
                case FtabValues.CLEAN:
                case FtabValues.MDETERM:
                case FtabValues.MINVERSE:
                case FtabValues.TERMINATE:
                case FtabValues.FACT:
                case FtabValues.ISNONTEXT:
                case FtabValues.ISLOGICAL:
                case FtabValues.DELETE_BAR:
                case FtabValues.UNREGISTER:
                case FtabValues.LENB:
                case FtabValues.ASC:
                case FtabValues.DBCS:
                case FtabValues.ELSE_IF:
                case FtabValues.SINH:
                case FtabValues.COSH:
                case FtabValues.TANH:
                case FtabValues.ASINH:
                case FtabValues.ACOSH:
                case FtabValues.ATANH:
                case FtabValues.INFO:
                case FtabValues.DELETE_TOOLBAR:
                case FtabValues.RESET_TOOLBAR:
                case FtabValues.EVALUATE:
                case FtabValues.ERROR_TYPE:
                case FtabValues.GAMMALN:
                case FtabValues.EVEN:
                case FtabValues.FISHER:
                case FtabValues.FISHERINV:
                case FtabValues.NORMSDIST:
                case FtabValues.NORMSINV:
                case FtabValues.ODD:
                case FtabValues.RADIANS:
                case FtabValues.DEGREES:
                case FtabValues.COUNTBLANK:
                case FtabValues.OPTIONS_LISTS_GET:
                case FtabValues.DATESTRING:
                case FtabValues.PHONETIC:
                case FtabValues.BAHTTEXT:
                case FtabValues.THAIDAYOFWEEK:
                case FtabValues.THAIDIGIT:
                case FtabValues.THAIMONTHOFYEAR:
                case FtabValues.THAINUMSOUND:
                case FtabValues.THAINUMSTRING:
                case FtabValues.THAISTRINGLENGTH:
                case FtabValues.ISTHAIDIGIT:
                case FtabValues.ROUNDBAHTDOWN:
                case FtabValues.ROUNDBAHTUP:
                case FtabValues.THAIYEAR:
                    return 1;
                case FtabValues.ARGUMENT:
                case FtabValues.SET_VALUE:
                case FtabValues.ROUND:
                case FtabValues.ABSREF:
                case FtabValues.REPT:
                case FtabValues.MOD:
                case FtabValues.TEXT:
                case FtabValues.RELREF:
                case FtabValues.ATAN2:
                case FtabValues.EXACT:
                case FtabValues.FREAD:
                case FtabValues.FWRITELN:
                case FtabValues.FWRITE:
                case FtabValues.MMULT:
                case FtabValues.INITIATE:
                case FtabValues.REQUEST:
                case FtabValues.EXECUTE:
                case FtabValues.ROUNDUP:
                case FtabValues.ROUNDDOWN:
                case FtabValues.FREQUENCY:
                case FtabValues.CHIDIST:
                case FtabValues.CHIINV:
                case FtabValues.COMBIN:
                case FtabValues.FLOOR:
                case FtabValues.CEILING:
                case FtabValues.PERMUT:
                case FtabValues.SUMXMY2:
                case FtabValues.SUMX2MY2:
                case FtabValues.SUMX2PY2:
                case FtabValues.CHITEST:
                case FtabValues.CORREL:
                case FtabValues.COVAR:
                case FtabValues.FTEST:
                case FtabValues.INTERCEPT:
                case FtabValues.PEARSON:
                case FtabValues.RSQ:
                case FtabValues.STEYX:
                case FtabValues.SLOPE:
                case FtabValues.LARGE:
                case FtabValues.SMALL:
                case FtabValues.QUARTILE:
                case FtabValues.PERCENTILE:
                case FtabValues.TRIMMEAN:
                case FtabValues.TINV:
                case FtabValues.POWER:
                case FtabValues.COUNTIF:
                case FtabValues.NUMBERSTRING:
                    return 2;
                case FtabValues.MID:
                case FtabValues.DCOUNT:
                case FtabValues.DSUM:
                case FtabValues.DAVERAGE:
                case FtabValues.DMIN:
                case FtabValues.DMAX:
                case FtabValues.DSTDEV:
                case FtabValues.DVAR:
                case FtabValues.DATE:
                case FtabValues.TIME:
                case FtabValues.SLN:
                case FtabValues.MIDB:
                case FtabValues.ENABLE_TOOL:
                case FtabValues.PRESS_TOOL:
                case FtabValues.CONFIDENCE:
                case FtabValues.CRITBINOM:
                case FtabValues.EXPONDIST:
                case FtabValues.FDIST:
                case FtabValues.FINV:
                case FtabValues.GAMMAINV:
                case FtabValues.LOGNORMDIST:
                case FtabValues.LOGINV:
                case FtabValues.NEGBINOMDIST:
                case FtabValues.NORMINV:
                case FtabValues.STANDARDIZE:
                case FtabValues.POISSON:
                case FtabValues.TDIST:
                case FtabValues.FORECAST:
                case FtabValues.DATEDIF:
                    return 3;
                case FtabValues.REPLACE:
                case FtabValues.SYD:
                case FtabValues.REPLACEB:
                case FtabValues.BINOMDIST:
                case FtabValues.GAMMADIST:
                case FtabValues.HYPGEOMDIST:
                case FtabValues.NORMDIST:
                case FtabValues.WEIBULL:
                case FtabValues.TTEST:
                case FtabValues.ISPMT:
                    return 4;
                default:
                    throw new NotImplementedException(
                        "Arg Count for " + value.ToString() + "not implemented.");
            }
        }

        public const PtgNumber ID = PtgNumber.PtgFunc;
        public ushort tab;

        public FtabValues Ftab
        {
            get
            {
                FtabValues val = (FtabValues) Enum.ToObject(typeof(FtabValues), tab);
                return val;
            }
        }

        public PtgFunc(FtabValues fTab, PtgDataType dt = PtgDataType.REFERENCE, bool setHighBit = false) : base(PtgNumber.PtgFunc, dt, setHighBit)
        {
            this.Data = "";
            this.Length = 3;
            this.popSize = 1;
            this.tab = Convert.ToUInt16(fTab);
            this.type = PtgType.Operator;
        }

        public PtgFunc(IStreamReader reader, PtgNumber ptgid)
            :
            base(reader, ptgid)
        {
            Debug.Assert(this.Id == ID);
            this.Length = 3;
            this.Data = "";
            this.type = PtgType.Operator;
             this.tab = this.Reader.ReadUInt16();
             
            this.popSize = 1;
        }
    }
}
