using ExcelReportGenerator.Converters.ExternalPropertiesConverters;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ExcelReportGenerator.Tests.Converters.ExternalPropertiesConverters
{
    [TestClass]
    public class RenderPriorityConverterTest
    {
        [TestMethod]
        public void TestConvert()
        {
            var converter = new RenderPriorityConverter();
            Assert.AreEqual(1, converter.Convert("1"));
            Assert.AreEqual(100, converter.Convert("100"));
            Assert.AreEqual(-10, converter.Convert("-10"));
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert("10.5"), "Cannot convert value \"10.5\" to int");
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert("BadValue"), "Cannot convert value \"BadValue\" to int");
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert(null), "RenderPriority property cannot be null or empty");
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert(string.Empty), "RenderPriority property cannot be null or empty");
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert(" "), "RenderPriority property cannot be null or empty");
        }
    }
}