using ExcelReporter.Helpers;
using ExcelReporter.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ExcelReporter.Tests.Helpers
{
    [TestClass]
    public class EnumHelperTest
    {
        [TestMethod]
        public void TestParse()
        {
            Assert.AreEqual(TestEnum.One, EnumHelper.Parse<TestEnum>("One"));
            ExceptionAssert.Throws<ArgumentException>(() => EnumHelper.Parse<TestEnum>("one", false));
            Assert.AreEqual(TestEnum.One, EnumHelper.Parse<TestEnum>("one"));
            Assert.AreEqual(TestEnum.Two, EnumHelper.Parse<TestEnum>("Two"));
            Assert.AreEqual(TestEnum.Three, EnumHelper.Parse<TestEnum>("Three"));
            ExceptionAssert.Throws<ArgumentException>(() => EnumHelper.Parse<TestEnum>("Four"));
            ExceptionAssert.Throws<ArgumentException>(() => EnumHelper.Parse<TestEnum>(null));
            ExceptionAssert.Throws<ArgumentException>(() => EnumHelper.Parse<int>("One"));
        }

        private enum TestEnum
        {
            One,
            Two,
            Three,
        }
    }
}