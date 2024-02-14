using System.Collections;
using System.Globalization;
using ExcelReportGenerator.Rendering;
using ExcelReportGenerator.Rendering.Panels.ExcelPanels;
using ExcelReportGenerator.Tests.CustomAsserts;
using ExcelReportGenerator.Tests.Rendering.Panels.ExcelPanels.PanelRenderTests;
using Assert = NUnit.Framework.Legacy.ClassicAssert;

namespace ExcelReportGenerator.Tests.Rendering;

public class SystemFunctionsTest
{
    [Test]
    public void TestGetDictVal()
    {
        var dict = new Dictionary<string, int>
        {
            ["Key1"] = 1,
            ["Key2"] = 2
        };

        Assert.AreEqual(1, SystemFunctions.GetDictVal(dict, "Key1"));
        Assert.AreEqual(2, SystemFunctions.GetDictVal(dict, "Key2"));

        ExceptionAssert.Throws<KeyNotFoundException>(() => SystemFunctions.GetDictVal(dict, "BadKey"),
            "The given key \"BadKey\" was not present in the dictionary");
        ExceptionAssert.Throws<ArgumentNullException>(() => SystemFunctions.GetDictVal(null, "Key1"));
        ExceptionAssert.Throws<ArgumentNullException>(() => SystemFunctions.GetDictVal(dict, null));
        ExceptionAssert.Throws<ArgumentException>(() => SystemFunctions.GetDictVal(Guid.NewGuid(), "Key1"),
            $"Parameter \"dictionary\" must implement {nameof(IDictionary)} interface");

        var objKey = new object();
        var objValue = new Random();
        var dict2 = new Hashtable
        {
            [1] = "One",
            ["Two"] = Guid.Empty,
            [objKey] = objValue
        };

        Assert.AreEqual("One", SystemFunctions.GetDictVal(dict2, 1));
        Assert.AreEqual(Guid.Empty, SystemFunctions.GetDictVal(dict2, "Two"));
        Assert.AreSame(objValue, SystemFunctions.GetDictVal(dict2, objKey));
    }

    [Test]
    public void TestTryGetDictVal()
    {
        var dict = new Dictionary<string, int>
        {
            ["Key1"] = 1,
            ["Key2"] = 2
        };

        Assert.AreEqual(1, SystemFunctions.TryGetDictVal(dict, "Key1"));
        Assert.AreEqual(2, SystemFunctions.TryGetDictVal(dict, "Key2"));

        Assert.IsNull(SystemFunctions.TryGetDictVal(dict, "BadKey"));
        Assert.IsNull(SystemFunctions.TryGetDictVal(null, "Key1"));
        Assert.IsNull(SystemFunctions.TryGetDictVal(dict, null));
        Assert.IsNull(SystemFunctions.TryGetDictVal(Guid.NewGuid(), "Key1"));

        var objKey = new object();
        var objValue = new Random();
        var dict2 = new Hashtable
        {
            [1] = "One",
            ["Two"] = Guid.Empty,
            [objKey] = objValue
        };

        Assert.AreEqual("One", SystemFunctions.TryGetDictVal(dict2, 1));
        Assert.AreEqual(Guid.Empty, SystemFunctions.TryGetDictVal(dict2, "Two"));
        Assert.AreSame(objValue, SystemFunctions.TryGetDictVal(dict2, objKey));
    }

    [Test]
    public void TestGetByIndex()
    {
        int[] intArray = {100, 200};
        Assert.AreEqual(100, SystemFunctions.GetByIndex(intArray, 0));
        Assert.AreEqual(200, SystemFunctions.GetByIndex(intArray, 1));

        var guid = Guid.NewGuid();
        var str = "Str";
        var rnd = new Random();
        object[] mixedArray = {guid, str, rnd};

        Assert.AreEqual(guid, SystemFunctions.GetByIndex(mixedArray, 0));
        Assert.AreEqual("Str", SystemFunctions.GetByIndex(mixedArray, 1));
        Assert.AreSame(rnd, SystemFunctions.GetByIndex(mixedArray, 2));

        IList<string> strList = new List<string> {"One", "Two"};
        Assert.AreEqual("One", SystemFunctions.GetByIndex(strList, 0));
        Assert.AreEqual("Two", SystemFunctions.GetByIndex(strList, 1));

        IList mixedList = new ArrayList();
        mixedList.Add(guid);
        mixedList.Add(str);
        mixedList.Add(rnd);

        Assert.AreEqual(guid, SystemFunctions.GetByIndex(mixedList, 0));
        Assert.AreEqual("Str", SystemFunctions.GetByIndex(mixedList, 1));
        Assert.AreSame(rnd, SystemFunctions.GetByIndex(mixedList, 2));

        ExceptionAssert.Throws<ArgumentNullException>(() => SystemFunctions.GetByIndex(null, 0));
        ExceptionAssert.Throws<ArgumentException>(() => SystemFunctions.GetByIndex(Guid.NewGuid(), 0),
            $"Parameter \"list\" must implement {nameof(IList)} interface");
        ExceptionAssert.Throws<IndexOutOfRangeException>(() => SystemFunctions.GetByIndex(mixedArray, -1));
        ExceptionAssert.Throws<IndexOutOfRangeException>(() => SystemFunctions.GetByIndex(mixedArray, 3));
    }

    [Test]
    public void TestTryGetByIndex()
    {
        int[] intArray = {100, 200};
        Assert.AreEqual(100, SystemFunctions.TryGetByIndex(intArray, 0));
        Assert.AreEqual(200, SystemFunctions.TryGetByIndex(intArray, 1));

        var guid = Guid.NewGuid();
        var str = "Str";
        var rnd = new Random();
        object[] mixedArray = {guid, str, rnd};

        Assert.AreEqual(guid, SystemFunctions.TryGetByIndex(mixedArray, 0));
        Assert.AreEqual("Str", SystemFunctions.TryGetByIndex(mixedArray, 1));
        Assert.AreSame(rnd, SystemFunctions.TryGetByIndex(mixedArray, 2));

        IList<string> strList = new List<string> {"One", "Two"};
        Assert.AreEqual("One", SystemFunctions.TryGetByIndex(strList, 0));
        Assert.AreEqual("Two", SystemFunctions.TryGetByIndex(strList, 1));

        IList mixedList = new ArrayList();
        mixedList.Add(guid);
        mixedList.Add(str);
        mixedList.Add(rnd);

        Assert.AreEqual(guid, SystemFunctions.TryGetByIndex(mixedList, 0));
        Assert.AreEqual("Str", SystemFunctions.TryGetByIndex(mixedList, 1));
        Assert.AreSame(rnd, SystemFunctions.TryGetByIndex(mixedList, 2));

        Assert.IsNull(SystemFunctions.TryGetByIndex(null, 0));
        Assert.IsNull(SystemFunctions.TryGetByIndex(Guid.NewGuid(), 0));
        Assert.IsNull(SystemFunctions.TryGetByIndex(mixedArray, -1));
        Assert.IsNull(SystemFunctions.TryGetByIndex(mixedArray, 3));
    }

    [Test]
    public void TestFormat()
    {
        Assert.AreEqual("31.01.2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "dd.MM.yyyy"));
        Assert.AreEqual("31.01.2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "dd.MM.yyyy", "RU"));
        Assert.AreEqual("31.01.2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "dd.MM.yyyy", "ru-ru"));
        Assert.AreEqual("31.01.2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "dd.MM.yyyy", 25)); // ru
        Assert.AreEqual("31.01.2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "dd.MM.yyyy", 1049)); // ru-RU
        Assert.AreEqual("31.01.2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "d"));
        Assert.AreEqual("01/31/2018",
            SystemFunctions.Format(new DateTime(2018, 1, 31), "d", CultureInfo.InvariantCulture));
        Assert.AreEqual("01/31/2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "d", 127)); // Invariant
        Assert.AreEqual("1/31/2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "d", "en"));
        Assert.AreEqual("1/31/2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "d", "en-US"));
        Assert.AreEqual("1/31/2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "d", 9)); // en
        Assert.AreEqual("1/31/2018", SystemFunctions.Format(new DateTime(2018, 1, 31), "d", 1033)); // en-US
        Assert.AreEqual(6535.676.ToString("0,0.##"), SystemFunctions.Format(6535.676, "0,0.##"));
        Assert.AreEqual(6535.676.ToString("0,0.##", CultureInfo.InvariantCulture),
            SystemFunctions.Format(6535.676, "0,0.##", CultureInfo.InvariantCulture));

        Assert.AreEqual(new DateTime(2018, 1, 31).ToString((string) null),
            SystemFunctions.Format(new DateTime(2018, 1, 31), null));
        Assert.AreEqual(new DateTime(2018, 1, 31).ToString(null, CultureInfo.InvariantCulture),
            SystemFunctions.Format(new DateTime(2018, 1, 31), null, CultureInfo.InvariantCulture));
        Assert.IsNull(SystemFunctions.Format(null, "dd.MM.yyyy"));
        Assert.IsNull(SystemFunctions.Format(null, "dd.MM.yyyy", CultureInfo.InvariantCulture));

        ExceptionAssert.Throws<ArgumentException>(() => SystemFunctions.Format(new Random(), "0,0.##"),
            $"Parameter \"input\" must implement {nameof(IFormattable)} interface");
        ExceptionAssert.Throws<CultureNotFoundException>(() =>
            SystemFunctions.Format(new DateTime(2018, 1, 31), "dd.MM.yyyy", "BadCulture"));
        ExceptionAssert.Throws<CultureNotFoundException>(() =>
            SystemFunctions.Format(new DateTime(2018, 1, 31), "dd.MM.yyyy", 10000000));
        ExceptionAssert.Throws<ArgumentOutOfRangeException>(() =>
            SystemFunctions.Format(new DateTime(2018, 1, 31), "dd.MM.yyyy", -1));
        ExceptionAssert.Throws<ArgumentException>(
            () => SystemFunctions.Format(new DateTime(2018, 1, 31), "dd.MM.yyyy", new Random()),
            "Invalid type \"Random\" of formatProvider");
    }

    [Test]
    public void TestFormatOnRender()
    {
        var report = new TestReport();
        var ws = report.Workbook.AddWorksheet("Test");
        var range = ws.Range(1, 1, 30, 30);

        ws.Cell(1, 1).Value = "''{sf:Format(p:FormatTest:Date, dd.MM.yyyy)}";
        ws.Cell(1, 2).Value = "''{sf:Format(p:FormatTest:Date, d, p:FormatTest:InvariantFormat)}";
        ws.Cell(1, 3).Value = "''{sf:Format(p:FormatTest:Date, d, p:FormatTest:UsFormat)}";
        ws.Cell(1, 4).Value = "''{sf:Format(p:FormatTest:Date, d, p:FormatTest:RuFormat)}";
        ws.Cell(1, 5).Value = "''{sf:Format(p:FormatTest:Date, d, RU)}";
        ws.Cell(1, 6).Value = "''{sf:Format(p:FormatTest:Date, d, [string]ru-RU)}";
        ws.Cell(1, 7).Value = "''{sf:Format(p:FormatTest:Date, d, en)}";
        ws.Cell(1, 8).Value = "''{sf:Format(p:FormatTest:Date, d, \"en-US\")}";
        ws.Cell(1, 9).Value = "''{sf:Format(p:FormatTest:Date, d, [int]25)}"; // ru
        ws.Cell(1, 10).Value = "''{sf:Format(p:FormatTest:Date, d, [int]1049)}"; // ru-RU
        ws.Cell(1, 11).Value = "''{sf:Format(p:FormatTest:Date, d, [int]9)}"; // en
        ws.Cell(1, 12).Value = "''{sf:Format(p:FormatTest:Date, d, [int]1033)}"; // en-US
        ws.Cell(1, 13).Value = "''{sf:Format(p:FormatTest:Date, d, [int]127)}"; // Invariant

        var panel = new ExcelPanel(range, report, report.TemplateProcessor);
        panel.Render();

        ExcelAssert.AreWorkbooksContentEquals(
            TestHelper.GetExpectedWorkbook(nameof(SystemFunctionsTest), "TestFormatOnRender"), ws.Workbook);

        //report.Workbook.SaveAs("test.xlsx");
    }

    private class FormatTest
    {
        public DateTime Date { get; set; } = new(2018, 1, 31);

        public IFormatProvider InvariantFormat { get; set; } = CultureInfo.InvariantCulture;

        public IFormatProvider UsFormat { get; set; } = new CultureInfo("en-US");

        public IFormatProvider RuFormat { get; set; } = new CultureInfo("ru-RU");
    }
}