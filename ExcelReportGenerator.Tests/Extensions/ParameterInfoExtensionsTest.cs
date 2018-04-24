using ExcelReportGenerator.Extensions;
using NUnit.Framework;
using System;
using System.Reflection;

namespace ExcelReportGenerator.Tests.Extensions
{
    public class ParameterInfoExtensionsTest
    {
        [Test]
        public void TestIsParams()
        {
            MethodInfo method = typeof(TestClass).GetMethod("Meth1");
            ParameterInfo[] parameters = method.GetParameters();
            Assert.IsFalse(parameters[0].IsParams());
            Assert.IsFalse(parameters[1].IsParams());
            Assert.IsTrue(parameters[2].IsParams());
        }

        [Test]
        public void TestHasDefaultValue()
        {
            MethodInfo method = typeof(TestClass).GetMethod("Meth2");
            ParameterInfo[] methodParams = method.GetParameters();
            Assert.IsFalse(methodParams[0].HasDefaultValue());
            Assert.IsTrue(methodParams[1].HasDefaultValue());
            Assert.IsTrue(methodParams[2].HasDefaultValue());
            Assert.IsTrue(methodParams[3].HasDefaultValue());
        }

        private class TestClass
        {
            public void Meth1(int arg1, string arg2, params string[] arg3)
            {
            }

            public void Meth2(int arg1, int arg2 = 0, DateTime? arg3 = null, object arg4 = null)
            {
            }
        }
    }
}