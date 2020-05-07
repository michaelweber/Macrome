

using b2xtranslator.StructuredStorage.Reader;

namespace b2xtranslator.Spreadsheet.XlsFileFormat.Structures
{
    /// <summary>
    /// An 8-byte union that specifies a floating-point value, or a non-numeric 
    /// value defined by the containing record. The type and meaning of the union 
    /// contents are determined by the most significant 2 bytes, and is defined in 
    /// the following table: 
    /// 
    ///   Value of most significant 2 bytes       Type and meaning of union contents
    ///     0xFFFF                                  A NilChartNum that specifies a non-numeric value, 
    ///                                             as defined by the containing record.
    ///                                             
    ///     Any other value.                        An Xnum that specifies a floating-point value.
    /// </summary>
    public class ChartNumNillable
    {
        /// <summary>
        /// The nullable Xnum value
        /// </summary>
        public double? value;

        public ChartNumNillable(IStreamReader reader)
        {
            //read the nullable double value 
            var b = reader.ReadBytes(8);
            if (b[6] == 0xFF && b[7] == 0xFF)
            {
                this.value = null;
            }
            else
            {
                this.value = System.BitConverter.ToDouble(b, 0);
            }
        }
    }
}
