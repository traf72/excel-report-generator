using ExcelReporter.Exceptions;
using ExcelReporter.Interfaces.Providers.DataItemValueProviders;
using System;

namespace ExcelReporter.Implementations.Providers.DataItemValueProviders
{
    public class HierarchicalDataItemValueProvider : IGenericDataItemValueProvider<HierarchicalDataItem>
    {
        private readonly IDataItemValueProviderFactory _factory;

        public HierarchicalDataItemValueProvider() : this(new DefaultDataItemValueProviderFactory())
        {
        }

        public HierarchicalDataItemValueProvider(IDataItemValueProviderFactory dataItemValueProviderFactory)
        {
            _factory = dataItemValueProviderFactory;
        }

        protected HierarchicalDataItem HierarchicalDataItem { get; private set; }

        public virtual object GetValue(string template, HierarchicalDataItem hierarchicalDataItem)
        {
            if (string.IsNullOrWhiteSpace(template))
            {
                throw new ArgumentException(Constants.EmptyStringParamMessage, nameof(template));
            }
            if (hierarchicalDataItem == null)
            {
                throw new ArgumentNullException(nameof(hierarchicalDataItem), Constants.NullParamMessage);
            }

            HierarchicalDataItem = hierarchicalDataItem;

            string dataItemTemplate;
            object dataItem = GetDataItemGivenHierarchy(template, out dataItemTemplate);
            return _factory.Create(dataItem)?.GetValue(dataItemTemplate, dataItem);
        }

        protected virtual object GetDataItemGivenHierarchy(string template, out string dataItemTemplate)
        {
            int lastColonIndex = template.LastIndexOf(":", StringComparison.Ordinal);
            if (lastColonIndex == -1)
            {
                dataItemTemplate = template;
                return HierarchicalDataItem.Value;
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
            dataItemTemplate = template.Substring(lastColonIndex + 1);
            return dataItem?.Value;
        }

        object IDataItemValueProvider.GetValue(string template, object hierarchicalDataItem)
        {
            return GetValue(template, (HierarchicalDataItem)hierarchicalDataItem);
        }
    }
}