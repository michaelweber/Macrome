using System.Diagnostics;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.Tools;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Records
{
    /// <summary>
    /// This record specifies additional extension properties of a date axis, 
    /// along with a CatSerRange record.
    /// </summary>
    [BiffRecord(RecordType.AxcExt)]
    public class AxcExt : BiffRecord
    {
        public const RecordType ID = RecordType.AxcExt;

        /// <summary>
        /// An unsigned integer that specifies the minimum date, as a date in the 
        /// date system specified by the Date1904 record, in the units defined by duBase. 
        /// 
        /// SHOULD <21> be less than or equal to catMax. If fAutoMin is set to 1, 
        /// MUST be ignored. If fDateAxis is set to 0, MUST be ignored.
        /// </summary>
        public ushort catMin;

        /// <summary>
        /// An unsigned integer that specifies the maximum date, as a date in the 
        /// date system specified by the Date1904 record, in the units defined by duBase. 
        /// SHOULD <22> be greater than or equal to catMin. If fAutoMax is set to 1, 
        /// MUST be ignored. If fDateAxis is set to 0, MUST be ignored.
        /// </summary>
        public ushort catMax;

        /// <summary>
        /// An unsigned integer that specifies the interval at which the major tick marks 
        /// are displayed on the axis, in the unit defined by duMajor. 
        /// 
        /// MUST be greater than or equal to catMinor when duMajor is equal to duMinor. 
        /// If fAutoMajor is set to 1, MUST be ignored. If fDateAxis is set to 0, MUST be ignored.
        /// </summary>
        public ushort catMajor;

        /// <summary>
        /// A DateUnit that specifies the unit of time to use for catMajor when 
        /// the axis is a date axis. If fDateAxis is set to 0, MUST be ignored.
        /// </summary>
        public DateUnit duMajor;

        /// <summary>
        /// An unsigned integer that specifies the interval at which the minor tick marks 
        /// are displayed on the axis, in a unit defined by duMinor. 
        /// 
        /// MUST be less than or equal to catMajor when duMajor is equal to duMinor. 
        /// If fAutoMinor is set to 1, MUST be ignored. If fDateAxis is set to 0, MUST be ignored.
        /// </summary>
        public ushort catMinor;

        /// <summary>
        /// A DateUnit that specifies the unit of time to use for catMinor when the 
        /// axis is a date axis. If fDateAxis is set to 0, MUST be ignored.
        /// </summary>
        public DateUnit duMinor;

        /// <summary>
        /// A DateUnit that specifies the smallest unit of time used by the axis. 
        /// 
        /// If fAutoBase is set to 1, this field MUST be ignored. If fDateAxis is set to 0, MUST be ignored.
        /// </summary>
        public DateUnit duBase;

        /// <summary>
        /// An unsigned integer that specifies at which date, as a date in the date system 
        /// specified by the Date1904 record, in the units defined by duBase, the value axis 
        /// crosses this axis. If fDateAxis is set to 0, MUST be ignored. 
        /// If fAutoCross is set to 1, MUST be ignored.
        /// </summary>
        public ushort catCrossDate;

        /// <summary>
        /// A bit that specifies whether catMin is calculated automatically. 
        /// 
        /// If fDateAxis is set to 0, MUST be ignored. 
        /// MUST be a value from the following table: 
        /// 
        ///     Value       Meaning
        ///     0           The value specified by catMin is used and catMin is not calculated automatically.
        ///     1           catMin is calculated such that the minimum data points value can be displayed.
        /// </summary>
        public bool fAutoMin;

        /// <summary>
        /// A A bit that specifies whether catMax is calculated automatically. 
        /// 
        /// If fDateAxis is set to 0, then fAutoMax MUST be ignored. If the value of the fMaxCross 
        /// field in the CatSerRange record is 1, then fAutoMax MUST be ignored. 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value       Meaning
        ///     0           The value specified by catMax is used and catMax is not calculated automatically.
        ///     1           catMax is calculated such that the minimum data points value can be displayed.
        /// </summary>
        public bool fAutoMax;

        /// <summary>
        /// A bit that specifies whether catMajor is calculated automatically. 
        /// 
        /// If fDateAxis is set to 0, MUST be ignored. 
        /// 
        ///     Value       Meaning
        ///     0           The value specified by catMajor is used and catMajor is not calculated automatically.
        ///     1           catMajor is calculated automatically.
        /// </summary>
        public bool fAutoMajor;

        /// <summary>
        /// A bit that specifies whether catMinor is calculated automatically. 
        /// 
        /// If fDateAxis is set to 0, MUST be ignored. 
        /// 
        ///     Value       Meaning
        ///     0           The value specified by catMinor is used and catMinor is not calculated automatically.
        ///     1           catMinor is calculated automatically.
        /// </summary>
        public bool fAutoMinor;

        /// <summary>
        /// A bit that specifies whether the axis is a date axis. 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value       Meaning
        ///     0           The axis is not a date axis.
        ///     1           The axis is a date axis.
        /// </summary>
        public bool fDateAxis;

        /// <summary>
        /// A bit that specifies whether the units of the date axis are chosen automatically. 
        /// 
        /// If fDateAxis is set to 0, MUST be ignored. 
        /// MUST be a value from the following table: 
        /// 
        ///     Value       Meaning
        ///     0           The value specified by duBase is used and duBase is not computed automatically.
        ///     1           duBase is calculated automatically.
        /// </summary>
        public bool fAutoBase;

        /// <summary>
        /// A bit that specifies whether catCrossDate is calculated automatically. 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value       Meaning
        ///     0           The value specified by catCrossDate is used and catCrossDate is not calculated automatically.
        ///     1           catCrossDate is calculated automatically such that it can be displayed.
        /// </summary>
        public bool fAutoCross;

        /// <summary>
        /// A bit that specifies whether the axis type is detected automatically. 
        /// 
        /// MUST be a value from the following table: 
        /// 
        ///     Value       Meaning
        ///     0           The axis will stay as specified by the fDateAxis field.
        ///     1           The axis will automatically become a date axis when the data it is related to contains date values; otherwise the axis will be a category axis.
        /// </summary>
        public bool fAutoDate;

        public AxcExt(IStreamReader reader, RecordType id, ushort length)
            : base(reader, id, length)
        {
            // assert that the correct record type is instantiated
            Debug.Assert(this.Id == ID);

            // initialize class members from stream
            this.catMin = reader.ReadUInt16();
            this.catMax = reader.ReadUInt16();
            this.catMajor = reader.ReadUInt16();
            this.duMajor = (DateUnit)reader.ReadUInt16();
            this.catMinor = reader.ReadUInt16();
            this.duMinor = (DateUnit)reader.ReadUInt16();
            this.duBase = (DateUnit)reader.ReadUInt16();
            this.catCrossDate = reader.ReadUInt16();

            ushort flags = reader.ReadUInt16();
            this.fAutoMin = Utils.BitmaskToBool(flags, 0x0001);
            this.fAutoMax = Utils.BitmaskToBool(flags, 0x0002);
            this.fAutoMajor = Utils.BitmaskToBool(flags, 0x0004);
            this.fAutoMinor = Utils.BitmaskToBool(flags, 0x0008);
            this.fDateAxis = Utils.BitmaskToBool(flags, 0x0010);
            this.fAutoBase = Utils.BitmaskToBool(flags, 0x0020);
            this.fAutoCross = Utils.BitmaskToBool(flags, 0x0040);
            this.fAutoDate = Utils.BitmaskToBool(flags, 0x0080);

            // assert that the correct number of bytes has been read from the stream
            Debug.Assert(this.Offset + this.Length == this.Reader.BaseStream.Position);
        }
    }
}
