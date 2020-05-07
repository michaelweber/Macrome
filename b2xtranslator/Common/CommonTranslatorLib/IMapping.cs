namespace b2xtranslator.CommonTranslatorLib
{
    public interface IMapping<T> where T : IVisitable
    {
        void Apply(T visited);
    }
}