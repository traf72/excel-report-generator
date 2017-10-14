using System;
using ReportEngine.Interfaces.DataSources;
using System.Collections.Generic;
using System.Linq;

namespace ReportEngine
{
    public class TestDataSource : IDataSource
    {
        private readonly IEnumerable<DataItem> _items = new List<DataItem>
        {
            new DataItem("Name_1", "Desc_1", 1.0m, 1),
            new DataItem("Name_2", "Desc_2", 2.0m, 1),
            new DataItem("Name_3", "Desc_3", 3.0m, 2),
            new DataItem("Name_4", "Desc_4", 4.0m, 2),
            new DataItem("Name_5", "Desc_5", 5.0m, 2),
            new DataItem("Name_6", "Desc_6", 6.0m, 3),
        };

        private readonly IEnumerable<DataItem2> _items2 = new List<DataItem2>
        {
            new DataItem2("Field1_1", "Field2_1"),
            new DataItem2("Field1_2", "Field2_2"),
            new DataItem2("Field1_3", "Field2_3"),
            new DataItem2("Field1_4", "Field2_4"),
        };

        public IEnumerable<int> GetGroups()
        {
            return new List<int> { 1, 2, 3 };
            //return new List<int> { 1, 2 };
            //return new List<int> { 1 };
        }

        public IEnumerable<DataItem> GetAllItems()
        {
            return _items;
            //return _items.Take(2);
        }

        public IEnumerable<DataItem> GetAllItemsPartial()
        {
            return _items.OrderBy(s => Guid.NewGuid()).Take(3);
        }

        public IEnumerable<DataItem> GetItemsByGroup(int groupId)
        {
            return _items.Where(i => i.Group == groupId);
        }

        public DataItem GetSingleItem(string name)
        {
            return _items.Single(i => i.Name == name);
        }

        public IEnumerable<DataItem> GetRandomItems()
        {
            Random rnd = new Random((int)DateTime.Now.Ticks & 0x0000FFFF);
            int count = rnd.Next(0, 7);
            return _items.Take(count);
            //return _items.Take(1);
        }

        public IEnumerable<DataItem2> GetRandomDataItems2()
        {
            Random rnd = new Random((int)DateTime.Now.Ticks & 0x0000FFFF);
            int count = rnd.Next(0, 5);
            return _items2.Take(count);
            //return _items2.Take(1);
        }
    }
}