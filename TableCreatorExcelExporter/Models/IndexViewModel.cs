namespace TableCreatorExcelExporter.Models
{
    public class IndexViewModel
    {
        public string TextData { get; set; }
        public List<string> Columns { get; set; } = new List<string>();
        public List<List<string>> Rows { get; set; } = new List<List<string>>();
    }
}
