namespace b2xtranslator.CommonTranslatorLib
{
    public interface IVisitable
    {
        void Convert<T>(T mapping);
    }
}