namespace WordDocumentProcessor.Models
{
    public class Report
    {
        public List<string> FilesProcessed { get; set; }
        public List<string> FilesWithMissingMetadata { get; set; }
        public int TotalWordCount { get; set; }

    }
}


