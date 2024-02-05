namespace MT.Models
{
    public class Title
    {
        public string name { get; set; }
        public string sheet { get; set; }
        public int tableCount { get; set; }

        public List<Column> columns { get; set; }
    }
}
