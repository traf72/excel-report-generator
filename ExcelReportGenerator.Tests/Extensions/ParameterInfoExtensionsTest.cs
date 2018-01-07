using System.Reflection;
using ExcelReportGenerator.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReportGenerator.Tests.Extensions
{
    [TestClass]
    public class ParameterInfoExtensionsTest
    {
        [TestMethod]
        public void TestIsParams()
        {
            MethodInfo method = typeof(TestClass).GetMethod("Meth1");
            ParameterInfo[] parameters = method.GetParameters();
            Assert.IsFalse(parameters[0].IsParams());
            Assert.IsFalse(parameters[1].IsParams());
            Assert.IsTrue(parameters[2].IsParams());
        }

        private class TestClass
        {
            public void Meth1(int arg1, string arg2, params string[] arg3)
            {
            }
        }
    }
}