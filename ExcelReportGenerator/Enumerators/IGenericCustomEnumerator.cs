namespace ExcelReportGenerator.Enumerators;

internal interface IGenericCustomEnumerator<out T> : ICustomEnumerator, IEnumerator<T>
{
}