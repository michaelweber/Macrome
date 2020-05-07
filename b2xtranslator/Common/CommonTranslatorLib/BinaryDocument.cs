namespace b2xtranslator.CommonTranslatorLib
{
    public abstract class BinaryDocument : IVisitable
    {
        #region IVisitable Members

        public abstract void Convert<T>(T mapping);
        
        #endregion
    }
}
