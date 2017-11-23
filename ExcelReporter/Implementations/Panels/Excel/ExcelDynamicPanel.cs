using ClosedXML.Excel;
using ExcelReporter.Interfaces.Panels.Excel;
using ExcelReporter.Interfaces.Reports;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using ExcelReporter.Enumerators;

namespace ExcelReporter.Implementations.Panels.Excel
{
    internal class ExcelDynamicPanel : ExcelNamedPanel
    {
        private readonly string _dataSourceTemplate;

        public ExcelDynamicPanel(string dataSourceTemplate, IXLNamedRange namedRange, IExcelReport report)
            : base(namedRange, report)
        {
            if (string.IsNullOrWhiteSpace(dataSourceTemplate))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(dataSourceTemplate));
            }
            _dataSourceTemplate = dataSourceTemplate;
        }

        public override void Render()
        {
            // Родительский контекст на данный тип панели никак влиять не может
            object data = Report.TemplateProcessor.GetValue(_dataSourceTemplate);
            IList<string> columns = new List<string> {"Id", "Name", "IsVip", "Description", "Type"};
            //IEnumerator enumerator = EnumeratorFactory.Create(data);
            //// Если данных нет, то просто удаляем сам шаблон
            //if (enumerator == null || !enumerator.MoveNext())
            //{
            //    DeletePanel(this);
            //    return;
            //}

            IXLCell firstCell = Range.FirstCell();
            IXLCell lastCell = firstCell;
            IXLCell currentCell = firstCell;
            currentCell.InsertCellsAfter(columns.Count - 1);
            foreach (string column in columns)
            {
                currentCell.Value = $"{{di:{column}}}";
                currentCell = currentCell.CellRight();
                lastCell = currentCell;
            }

            IXLWorksheet ws = Range.Worksheet;
            IXLRange range = ws.Range(firstCell, lastCell);
            string name = $"{range}_{Guid.NewGuid():N}";
            range.AddToNamed(name, XLScope.Worksheet);

            var dataSourcePanel = new ExcelDataSourcePanel(_dataSourceTemplate, ws.NamedRange(name), Report)
            {
                ShiftType = ShiftType,
                AfterRenderMethodName = AfterRenderMethodName,
                BeforeRenderMethodName = BeforeRenderMethodName,
                Type = Type,
            };

            dataSourcePanel.Render();
        }

        protected override IExcelPanel CopyPanel(IXLCell cell)
        {
            var panel = new ExcelDynamicPanel(_dataSourceTemplate, CopyNamedRange(cell), Report);
            FillCopyProperties(panel);
            return panel;
        }

        //private IList<string> GetDataColumns(object data)
        //{
        //    switch (data)
        //    {
        //        case null:
        //            return null;
        //        case IDataReader dr:
                    
        //        //case DataTable dt:
        //        //    return dt.AsEnumerable().GetEnumerator();
        //        //case DataSet ds:
        //        //    return new DataSetEnumerator(ds);
        //        //case IEnumerable e:
        //        //    return e.GetEnumerator();
        //    }

        //    return null;
        //}
    }
}