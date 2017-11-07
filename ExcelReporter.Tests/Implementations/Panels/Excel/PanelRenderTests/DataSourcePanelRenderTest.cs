using ClosedXML.Excel;
using ExcelReporter.Implementations.Panels.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace ExcelReporter.Tests.Implementations.Panels.Excel.PanelRenderTests
{
    [TestClass]
    public class DataSourcePanelRenderTest
    {
        [TestMethod]
        public void TestPanelRender()
        {
            var report = new TestReport();
            IXLWorksheet ws = report.Workbook.AddWorksheet("Test");
            IXLRange range = ws.Range(1, 1, 2, 4);
            range.AddToNamed("TestRange", XLScope.Worksheet);

            ws.Cell(1, 1).Value = "{di:Name}";
            ws.Cell(1, 2).Value = "{di:Date}";
            ws.Cell(1, 3).Value = "{di:Sum}";
            ws.Cell(1, 4).Value = "{di:Contacts}";
            ws.Cell(2, 1).Value = "{di:Contacts.Phone}";
            ws.Cell(2, 2).Value = "{di:Contacts.Fax}";

            ws.Cell(1, 5).Value = "{di:Name}";
            ws.Cell(3, 1).Value = "{di:Date}";

            var panel = new ExcelDataSourcePanel("m:TestDataProvider:GetSingle()", ws.NamedRange("TestRange"), report);
            panel.Render();

            Assert.AreEqual(8, ws.CellsUsed().Count());
            Assert.AreEqual("Test", ws.Cell(1, 1).Value);
            Assert.AreEqual(new DateTime(2017, 11, 1), ws.Cell(1, 2).Value);
            Assert.AreEqual(55.76, ws.Cell(1, 3).Value);
            Assert.AreEqual("15_345", ws.Cell(1, 4).Value);
            Assert.AreEqual(15d, ws.Cell(2, 1).Value);
            Assert.AreEqual(345d, ws.Cell(2, 2).Value);

            Assert.AreEqual("{di:Name}", ws.Cell(1, 5).Value);
            Assert.AreEqual("{di:Date}", ws.Cell(3, 1).Value);

            Assert.AreEqual(0, ws.NamedRanges.Count());
            Assert.AreEqual(0, ws.Workbook.NamedRanges.Count());

            Assert.AreEqual(1, ws.Workbook.Worksheets.Count);

            report.Workbook.SaveAs("test.xlsx");
        }

        private class TestItem
        {
            public TestItem(string name, DateTime date, decimal sum, Contacts contacts = null)
            {
                Name = name;
                Date = date;
                Sum = sum;
                Contacts = contacts;
            }

            public string Name { get; set; }

            public DateTime Date { get; set; }

            public decimal Sum { get; set; }

            public Contacts Contacts { get; set; }
        }

        private class Contacts
        {
            public Contacts(string phone, string fax)
            {
                Phone = phone;
                Fax = fax;
            }

            public string Phone { get; set; }

            public string Fax { get; set; }

            public override string ToString()
            {
                return $"{Phone}_{Fax}";
            }
        }

        private class TestDataProvider
        {
            public TestItem GetSingle()
            {
                return new TestItem("Test", new DateTime(2017, 11, 1), 55.76m, new Contacts("15", "345"));
            }
        }
    }
}