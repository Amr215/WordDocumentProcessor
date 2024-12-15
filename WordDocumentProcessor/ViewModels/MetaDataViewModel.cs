namespace WordDocumentProcessor.ViewModels
{
    public class MetaDataViewModel
    {
        public string FileName { get; set; }
        public string? Title { get; set; }
        public string? Author { get; set; }
        public DateTime? CreationDate { get; set; }
        public int? WordCount { get; set; }
        public int? PageCount { get; set; }
    }
}


