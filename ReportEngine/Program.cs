namespace ReportEngine
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            new TestReport().Run();
            new TestReport2().Run();
            new TestReport3().Run();
        }
    }
}