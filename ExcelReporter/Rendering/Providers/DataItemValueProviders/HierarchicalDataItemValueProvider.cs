using System;
using ExcelReporter.Exceptions;
using ExcelReporter.Helpers;

namespace ExcelReporter.Rendering.Providers.DataItemValueProviders
{
    /// <summary>
    /// Provides values from hierarchical data item
    /// </summary>
    public class HierarchicalDataItemValueProvider : IGenericDataItemValueProvider<HierarchicalDataItem>
    {
        private readonly IDataItemValueProviderFactory _factory;

        public HierarchicalDataItemValueProvider() : this(new DefaultDataItemValueProviderFactory())
        {
        }

        internal HierarchicalDataItemValueProvider(IDataItemValueProviderFactory dataItemValueProviderFactory)
        {
            _factory = dataItemValueProviderFactory;
        }

        protected HierarchicalDataItem HierarchicalDataItem { get; private set; }

        /// <summary>
        /// Returns value from hierarchical data item based on template
        /// </summary>
        public virtual object GetValue(string template, HierarchicalDataItem hierarchicalDataItem)
        {
            if (string.IsNullOrWhiteSpace(template))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(template));
            }
            HierarchicalDataItem = hierarchicalDataItem ?? throw new ArgumentNullException(nameof(hierarchicalDataItem), ArgumentHelper.NullParamMessage);

            var (dataItem, dataItemTemplate) = GetDataItemGivenHierarchy(template);
            return _factory.Create(dataItem)?.GetValue(dataItemTemplate, dataItem);
        }

        /// <summary>
        /// Returns real data item object given hierarchy and template for this data item based on input template
        /// </summary>
        protected virtual (object dataItem, string dataItemTemplate) GetDataItemGivenHierarchy(string template)
        {
            int lastColonIndex = template.LastIndexOf(":", StringComparison.Ordinal);
            if (lastColonIndex == -1)
            {
                return (HierarchicalDataItem.Value, template);
            }

            string[] parentTemplateParts = template.Substring(0, lastColonIndex).Split(':');
            HierarchicalDataItem dataItem = HierarchicalDataItem;
            foreach (string part in parentTemplateParts)
            {
                if (!part.Trim().Equals("parent", StringComparison.OrdinalIgnoreCase))
                {
                    throw new IncorrectTemplateException($"Template \"{template}\" is incorrect");
                }
                dataItem = dataItem.Parent;
            }
            return (dataItem?.Value, template.Substring(lastColonIndex + 1).Trim());
        }

        object IDataItemValueProvider.GetValue(string template, object hierarchicalDataItem)
        {
            return GetValue(template, (HierarchicalDataItem)hierarchicalDataItem);
        }
    }
}