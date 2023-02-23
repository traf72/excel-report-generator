using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Rendering.Providers.VariableProviders;
using ExcelReportGenerator.Tests.CustomAsserts;

namespace ExcelReportGenerator.Tests.Rendering.Providers.VariableProviders;

public class DefaultVariableProviderTest
{
    [Test]
    public void TestGetVariable()
    {
        SystemVariableProvider variableProvider = new CustomVariableProvider
        {
            RenderDate = new DateTime(2018, 1, 1),
            CustomProp = 999,
            SheetName = "Sheet",
            SheetNumber = 3
        };

        Assert.AreEqual(new DateTime(2018, 1, 1), variableProvider.GetVariable("RenderDate"));
        Assert.AreEqual(999, variableProvider.GetVariable("CustomProp"));
        Assert.AreEqual("Sheet", variableProvider.GetVariable("SheetName"));
        Assert.AreEqual(3, variableProvider.GetVariable("SheetNumber"));
        ExceptionAssert.Throws<InvalidVariableException>(() => variableProvider.GetVariable("BadVariable"),
            $"Cannot find variable with name \"BadVariable\" in class {typeof(CustomVariableProvider).Name} and all its parents. Variable must be public instance property.");

        ExceptionAssert.Throws<ArgumentException>(() => variableProvider.GetVariable(null));
        ExceptionAssert.Throws<ArgumentException>(() => variableProvider.GetVariable(string.Empty));
        ExceptionAssert.Throws<ArgumentException>(() => variableProvider.GetVariable(" "));
    }

    private class CustomVariableProvider : SystemVariableProvider
    {
        public new string SheetName { get; set; }

        public int CustomProp { get; set; }
    }
}