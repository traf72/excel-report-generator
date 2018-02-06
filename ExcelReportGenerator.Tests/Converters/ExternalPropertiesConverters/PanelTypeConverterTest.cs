using ExcelReportGenerator.Converters.ExternalPropertiesConverters;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Tests.CustomAsserts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace ExcelReportGenerator.Tests.Converters.ExternalPropertiesConverters
{
    [TestClass]
    public class PanelTypeConverterTest
    {
        [TestMethod]
        public void TestConvert()
        {
            var converter = new PanelTypeConverter();
            Assert.AreEqual(PanelType.Horizontal, converter.Convert("Horizontal"));
            Assert.AreEqual(PanelType.Horizontal, converter.Convert("horizontal"));
            Assert.AreEqual(PanelType.Vertical, converter.Convert("Vertical"));
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert("BadValue"), "Value \"BadValue\" is invalid for Type property");
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert(null), "Type property cannot be null or empty");
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert(string.Empty), "Type property cannot be null or empty");
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert(" "), "Type property cannot be null or empty");
        }
    }
}