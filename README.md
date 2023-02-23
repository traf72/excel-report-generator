# ExcelReportGenerator

This library allows you to render data to Microsoft Excel by marking Excel sheets using panels and templates. It makes it easy to connect to various data sources like IDataReader, DataSet, DataTable, IEnumerable<T> and others and to render data from them to Excel. You can also apply aggregation, various types of grouping, formatting and so on to the data. You can include the library to your project via [NuGet package](https://www.nuget.org/packages/ExcelReportGenerator).

## How to use?
First of all you have to create a report template in Microsoft Excel. You need to mark up an excel sheet (or multiple sheets) using panels and templates which will be handled by the library. You can add any other formatting you want. All styles will be preserved after rendering. Your template can look like this:

**Template**

![template](https://user-images.githubusercontent.com/45209977/48908282-e29a8b80-ee7a-11e8-9b54-6997ba617474.png)

Next you need to create a report class that will supply data for this template

**Code**

```c#
    public class ReportSample
    {
        private readonly DataProvider _dataProvider = new DataProvider();

        public string ReportName => "Grouping with Panel Hierarchy";

        public IEnumerable<string> GetDepartments()
        {
            return GetAllEmployees().Select(e => e.DepartmentName).Distinct();
        }

        public IEnumerable<DataProvider.Result> GetDepartmentEmployees(string department)
        {
            return GetAllEmployees().Where(e => e.DepartmentName == department).ToArray();
        }

        public IEnumerable<DataProvider.Result> GetAllEmployees()
        {
            return _dataProvider.GetEmployeesAsIEnumerable().ToArray();
        }

        public string ConvertGender(string gender)
        {
            return gender == "M" ? "Male" : "Female";
        }
    }
```

**Result**

![result](https://user-images.githubusercontent.com/45209977/48908531-94d25300-ee7b-11e8-8022-5c6cfdfca4e3.png)

The detailed documentation is inside the Docs folder. For more information see also ExcelReportGenerator.Samples.
