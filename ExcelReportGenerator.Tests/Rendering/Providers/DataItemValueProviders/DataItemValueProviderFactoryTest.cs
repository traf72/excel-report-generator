using System.Collections.Generic;
using System.Data;
using ExcelReportGenerator.Rendering.Providers.DataItemValueProviders;
using NUnit.Framework;
using NSubstitute;

namespace ExcelReportGenerator.Tests.Rendering.Providers.DataItemValueProviders
{
    
    public class DataItemValueProviderFactoryTest
    {
        [Test]
        public void TestCreate()
        {
            var factory = new DataItemValueProviderFactory();
            Assert.IsInstanceOf<ObjectPropertyValueProvider>(factory.Create(null));
            Assert.IsInstanceOf<DictionaryValueProvider<object>>(factory.Create(new Dictionary<string, object>()));
            Assert.IsInstanceOf<DictionaryValueProvider<int>>(factory.Create(new Dictionary<string, int>()));
            Assert.IsInstanceOf<DictionaryValueProvider<string>>(factory.Create(new Dictionary<string, string>()));

            var dataTable = new DataTable();
            dataTable.Columns.Add("Column", typeof(int));
            dataTable.Rows.Add(1);
            Assert.IsInstanceOf<DataRowValueProvider>(factory.Create(dataTable.Rows[0]));

            Assert.IsInstanceOf<DataReaderValueProvider>(factory.Create(Substitute.For<IDataReader>()));

            Assert.IsInstanceOf<ObjectPropertyValueProvider>(factory.Create(new int()));
            Assert.IsInstanceOf<ObjectPropertyValueProvider>(factory.Create(new object()));
            Assert.IsInstanceOf<ObjectPropertyValueProvider>(factory.Create(new Dictionary<object, string>()));
        }
    }
}