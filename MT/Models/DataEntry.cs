namespace MT.Models
{
    public class DataEntry
    {
        public object? value { get; set; }
        public int i { get; set; }
        public int j { get; set; }
        public string cell { get; set; }
        public int tableCount { get; set; }
        public Dictionary<string, int[]> datasource { get; set; }
        public string SheetName { get; set; }
    }
}
