using ExcelReportGenerator.Exceptions;
using ExcelReportGenerator.Helpers;
using System;
using ExcelReportGenerator.Attributes;

namespace ExcelReportGenerator.Rendering.Providers.DataItemValueProviders
{
    /// <summary>
    /// Default implementation of <see cref="IGenericDataItemValueProvider{T}" />
    /// Provides values from hierarchical data item
    /// </summary>
    [LicenceKeyPart(U = true)]
    public class DefaultDataItemValueProvider : IGenericDataItemValueProvider<HierarchicalDataItem>
    {
        private readonly IDataItemValueProviderFactory _factory;

        public DefaultDataItemValueProvider() : this(new DataItemValueProviderFactory())
        {
        }

        internal DefaultDataItemValueProvider(IDataItemValueProviderFactory dataItemValueProviderFactory)
        {
            _factory = dataItemValueProviderFactory;
        }

        // Get or set the template if you want to return the data item itself
        internal string DataItemSelfTemplate { get; set; }


        /// <inheritdoc />
        /// <exception cref="ArgumentException">Thrown when <paramref name="template" /> is null, empty string or whitespace</exception>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="hierarchicalDataItem"/> is null</exception>
        /// <exception cref="InvalidTemplateException"></exception>
        public virtual object GetValue(string template, HierarchicalDataItem hierarchicalDataItem)
        {
            if (string.IsNullOrWhiteSpace(template))
            {
                throw new ArgumentException(ArgumentHelper.EmptyStringParamMessage, nameof(template));
            }
            if (hierarchicalDataItem == null)
            {
                throw new ArgumentNullException(nameof(hierarchicalDataItem), ArgumentHelper.NullParamMessage);
            }

            var (dataItem, dataItemTemplate) = GetDataItemGivenHierarchy(template, hierarchicalDataItem);
            if (dataItemTemplate == DataItemSelfTemplate)
            {
                return dataItem;
            }

            return _factory.Create(dataItem)?.GetValue(dataItemTemplate, dataItem);
        }

        // Returns real data item object given hierarchy and template for this data item based on input template
        private (object dataItem, string dataItemTemplate) GetDataItemGivenHierarchy(string template, HierarchicalDataItem hierarchicalDataItem)
        {
            int lastColonIndex = template.LastIndexOf(":", StringComparison.Ordinal);
            if (lastColonIndex == -1)
            {
                return (hierarchicalDataItem.Value, template);
            }

            string[] parentTemplateParts = template.Substring(0, lastColonIndex).Split(':');
            HierarchicalDataItem dataItem = hierarchicalDataItem;
            foreach (string part in parentTemplateParts)
            {
                if (!part.Trim().Equals("parent", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidTemplateException($"Template \"{template}\" is invalid");
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