using ExcelReporter.Implementations.Providers;
using ExcelReporter.Interfaces.Providers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Reflection;

namespace ExcelReporter.Tests.Implementations.Providers
{
    [TestClass]
    public class TypeInstanceProviderTest
    {
        [TestMethod]
        public void TestGetInstance()
        {
            MyAssert.Throws<ArgumentNullException>(() => new TypeInstanceProvider(null));

            ITypeProvider typeProvider = new TypeProvider(GetType(), Assembly.GetExecutingAssembly());
            ITypeInstanceProvider typeInstanceProvider = new TypeInstanceProvider(typeProvider);

            Assert.AreEqual(GetType().FullName, typeInstanceProvider.GetInstance(null).GetType().FullName);
            Assert.AreEqual(GetType().FullName, typeInstanceProvider.GetInstance(string.Empty).GetType().FullName);
            Assert.AreEqual(GetType().FullName, typeInstanceProvider.GetInstance(" ").GetType().FullName);

            Assert.AreEqual(GetType().FullName, typeInstanceProvider.GetInstance("TypeInstanceProviderTest").GetType().FullName);
            Assert.AreEqual(GetType().FullName, typeInstanceProvider.GetInstance("ExcelReporter.Tests.Implementations.Providers:TypeInstanceProviderTest").GetType().FullName);
        }
    }
}