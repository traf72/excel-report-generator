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
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert("BadValue"), "Value \"BadValue\" is invalid for PanelType property");
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert(null), "PanelType property cannot be null or empty");
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert(string.Empty), "PanelType property cannot be null or empty");
            ExceptionAssert.Throws<ArgumentException>(() => converter.Convert(" "), "PanelType property cannot be null or empty");
        }
    }
}