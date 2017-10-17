namespace ExcelReporter.Interfaces.Providers
{
    public interface ITypeInstanceProvider
    {
        object GetInstance(string typeTemplate);
    }
}