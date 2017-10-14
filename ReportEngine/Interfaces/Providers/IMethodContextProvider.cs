namespace ReportEngine.Interfaces.Providers
{
    public interface IMethodContextProvider
    {
        object GetMethodContext(string className);
    }
}