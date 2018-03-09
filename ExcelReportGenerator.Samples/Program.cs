using System;
using System.Windows.Forms;
using SimpleInjector;

namespace ExcelReportGenerator.Samples
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            Ioc.Container = new Container();
            Ioc.Container.Register<DataProvider>(Lifestyle.Singleton);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SamplesForm());
        }
    }
}