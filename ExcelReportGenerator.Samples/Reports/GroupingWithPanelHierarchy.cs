namespace ExcelReportGenerator.Samples.Reports;

public class GroupingWithPanelHierarchy : ReportBase
{
    private readonly DataProvider _dataProvider = new DataProvider();

    private DataProvider.Result[] _allEmployeesCache;

    private readonly IDictionary<string, DataProvider.Result[]> _employeesByDepartmentCache = new Dictionary<string, DataProvider.Result[]>();

    public override string ReportName => "Grouping with Panel Hierarchy";

    public IEnumerable<string> GetDepartments()
    {
        return GetAllEmployees().Select(e => e.DepartmentName).Distinct();
    }

    public IEnumerable<DataProvider.Result> GetDepartmentEmployees(string department)
    {
        if (_employeesByDepartmentCache.TryGetValue(department, out var result))
        {
            return result;
        }

        result = GetAllEmployees().Where(e => e.DepartmentName == department).ToArray();
        _employeesByDepartmentCache[department] = result;
        return result;
    }

    public IEnumerable<DataProvider.Result> GetAllEmployees()
    {
        return _allEmployeesCache ??= _dataProvider.GetEmployeesAsIEnumerable().ToArray();
    }
}