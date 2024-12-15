using Microsoft.AspNetCore.Mvc;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDocumentProcessor.Models;
using WordDocumentProcessor.ViewModels;

namespace WordDocumentProcessor.Controllers
{
    public class WordController : Controller
    {
        [HttpGet]
        public IActionResult ProcessForm()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ProcessDocuments(List<IFormFile> files, string outputPath)
        {
            // Set default output path if not provided
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                outputPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "Reports", "word_Report.json");
            }
            else if (!outputPath.EndsWith("word_Report.json"))
            {
                outputPath = Path.Combine(outputPath, "word_Report.json");
            }

            // Ensure the Reports directory exists
            CreateDirectoryIfNotExists(Path.GetDirectoryName(outputPath));

            var report = new Report
            {
                FilesProcessed = new List<string>(),
                FilesWithMissingMetadata = new List<string>(),
                TotalWordCount = 0
            };

            List<MetaDataViewModel> listMetaData = new List<MetaDataViewModel>();

            foreach (var file in files)
            {
                // Ensure files are saved to wwwroot/UploadedFiles
                var uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "UploadedFiles");
                CreateDirectoryIfNotExists(uploadsFolder);

                var filePath = Path.Combine(uploadsFolder, file.FileName);
                await using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await file.CopyToAsync(stream);
                }

                if (System.IO.File.Exists(filePath))
                {
                    var metadata = ExtractMetadata(filePath);
                    var wordCount = CountWords(filePath);
                    var pageCount = GetPageCount(filePath);

                    report.FilesProcessed.Add(file.FileName);
                    report.TotalWordCount += wordCount;

                    if (metadata == null ||
                        string.IsNullOrWhiteSpace(metadata.Title) ||
                        string.IsNullOrWhiteSpace(metadata.Author) ||
                        metadata.CreationDate == null)
                    {
                        report.FilesWithMissingMetadata.Add(file.FileName);
                    }

                    MetaDataViewModel metaDataViewModel = new MetaDataViewModel
                    {
                        FileName = file.FileName,
                        Title = metadata?.Title,
                        Author = metadata?.Author,
                        CreationDate = metadata?.CreationDate,
                        PageCount = pageCount,
                        WordCount = wordCount,
                    };
                    listMetaData.Add(metaDataViewModel);
                }
            }

            // Serialize report and write to file
            string jsonReport = JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true });
            await System.IO.File.WriteAllTextAsync(outputPath, jsonReport);

            // Prepare ViewModel
            ResultViewModel resultViewModel = new ResultViewModel
            {
                Report = report,
                ListMetadata = listMetaData
            };

            return View("Result", resultViewModel);
        }

        private Metadata ExtractMetadata(string filePath)
        {
            try
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false))
                {
                    var properties = wordDocument.PackageProperties;
                    var title = properties.Title;
                    var author = properties.Creator;
                    var creationDate = properties.Created;

                    return new Metadata
                    {
                        Title = title,
                        Author = author,
                        CreationDate = creationDate
                    };
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading file {filePath}: {ex.Message}");
                return null;
            }
        }

        private int CountWords(string filePath)
        {
            try
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false))
                {
                    var body = wordDocument.MainDocumentPart.Document.Body;
                    var textElements = body.Descendants<Text>();

                    return textElements.Sum(te => te.Text.Split(' ').Length);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error counting words in file {filePath}: {ex.Message}");
                return 0;
            }
        }

        private int GetPageCount(string filePath)
        {
            try
            {
                var doc = new Aspose.Words.Document(filePath);
                return doc.PageCount;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting page count for {filePath}: {ex.Message}");
                return 0;
            }
        }

        private void CreateDirectoryIfNotExists(string directoryPath)
        {
            if (!string.IsNullOrEmpty(directoryPath) && !Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }
        }
    }
}
