namespace b2xtranslator.Tools
{
    public class TwipsValue
    {
        /// <summary>
        /// The dots per inch value that should be used.
        /// </summary>
        public const double Dpi = 72.0;

        private double value;

        /// <summary>
        /// Creates a new TwipsValue for the given value.
        /// </summary>
        /// <param name="value"></param>
        public TwipsValue(double value)
        {
            this.value = value;
        }

        /// <summary>
        /// Converts the twips to pt
        /// </summary>
        /// <returns></returns>
        public double ToPoints()
        {
            return this.value / 20;
        }

        /// <summary>
        /// Converts the twips to inch
        /// </summary>
        /// <returns></returns>
        public double ToInch()
        {
            return this.value / (TwipsValue.Dpi * 20);
        }

        /// <summary>
        /// Converts the twips to mm
        /// </summary>
        /// <returns></returns>
        public double ToMm()
        {
            return ToInch() * 25.399931;
        }

        /// <summary>
        /// Converts the twips to cm
        /// </summary>
        /// <returns></returns>
        public double ToCm()
        {
            return ToMm() / 10.0;
        }
    }
}
