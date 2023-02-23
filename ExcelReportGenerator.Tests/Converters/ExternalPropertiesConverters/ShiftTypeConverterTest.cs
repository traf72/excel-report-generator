using ExcelReportGenerator.Converters.ExternalPropertiesConverters;
using ExcelReportGenerator.Enums;
using ExcelReportGenerator.Tests.CustomAsserts;

namespace ExcelReportGenerator.Tests.Converters.ExternalPropertiesConverters;

public class ShiftTypeConverterTest
{
    [Test]
    public void TestConvert()
    {
        var converter = new ShiftTypeConverter();
        Assert.AreEqual(ShiftType.Cells, converter.Convert("Cells"));
        Assert.AreEqual(ShiftType.Cells, converter.Convert("cells"));
        Assert.AreEqual(ShiftType.Row, converter.Convert("Row"));
        Assert.AreEqual(ShiftType.NoShift, converter.Convert("NoShift"));
        ExceptionAssert.Throws<ArgumentException>(() => converter.Convert("BadValue"),
            "Value \"BadValue\" is invalid for ShiftType property");
        ExceptionAssert.Throws<ArgumentException>(() => converter.Convert(null),
            "ShiftType property cannot be null or empty");
        ExceptionAssert.Throws<ArgumentException>(() => converter.Convert(string.Empty),
            "ShiftType property cannot be null or empty");
        ExceptionAssert.Throws<ArgumentException>(() => converter.Convert(" "),
            "ShiftType property cannot be null or empty");
    }
}