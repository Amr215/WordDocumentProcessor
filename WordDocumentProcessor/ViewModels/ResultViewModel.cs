using WordDocumentProcessor.Models;

namespace WordDocumentProcessor.ViewModels
{
    public class ResultViewModel
    {
        public Report Report { get; set; }
        public List<MetaDataViewModel> ListMetadata { get; set; } = new List<MetaDataViewModel>();
    }
}

