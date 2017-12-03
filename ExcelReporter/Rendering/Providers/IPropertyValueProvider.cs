namespace ExcelReporter.Rendering.Providers
{
    public interface IPropertyValueProvider
    {
        /// <summary>
        /// Provides property value base on string property template
        /// </summary>
        object GetValue(string propertyTemplate);
    }
}