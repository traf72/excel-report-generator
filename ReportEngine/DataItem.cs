namespace ReportEngine
{
    public class DataItem
    {
        public DataItem(string name, string description, decimal value, int group)
        {
            Name = name;
            Description = description;
            Value = value;
            Group = group;
        }

        public string Name { get; set; }

        public string Description { get; set; }

        public decimal Value { get; set; }

        public int Group { get; set; }
    }
}