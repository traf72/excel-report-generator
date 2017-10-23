namespace ExcelReporter.Iterators
{
    public interface IIterator<out T>
    {
        T Next();

        bool HaxNext();

        void Reset();
    }
}