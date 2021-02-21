namespace SimMatch
{
    using System.Collections.Generic;
    using System.Data;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;

    using Annytab.Stemmer;

    using CsvHelper;
    using CsvHelper.Configuration;
    using CsvHelper.Excel;

    using ExcelDataReader;

    using StopWord;

    public class FileService
    {
        public DataSet GetRecords(string file)
        {
            using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            // reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true
                        }
                    });

                    // The result of each spreadsheet is in result.Tables
                    return result;
                }
            }
        }

        public IEnumerable<T> GetRecords<T>(string file, bool isExcel = false, string sheetName = null)
        {
            if (isExcel)
            {
                var excelParser = new ExcelParser(file, sheetName, new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    PrepareHeaderForMatch = (header, i) => header.SanitizeHeader()
                });

                using (var csvReader = new CsvReader(excelParser))
                {
                    var results = csvReader.GetRecords<T>().ToList();
                    return results;
                }
            }
            else
            {
                var textReader = new StreamReader(file);
                using (var csvReader = new CsvReader(textReader, new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    PrepareHeaderForMatch = (header, i) => header.SanitizeHeader()
                }))
                {
                    var results = csvReader.GetRecords<T>().ToList();
                    return results;
                }
            }
        }

        public void WriteRecords<T>(IEnumerable<T> records, string outputPath)
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            using (var excelWriter = new ExcelWriter(outputPath, CultureInfo.InvariantCulture))
            {
                excelWriter.WriteRecords(records);
            }
        }
    }

    public static class SimMatchExtensions
    {
        private static readonly EnglishStemmer Stemmer = new EnglishStemmer();
        private static readonly List<string> StopWordsList = StopWords.GetStopWords("en").ToList();

        public static string SanitizeHeader(this string header)
        {
            return Regex.Replace(header, @"(\s+|@|&|'|\(|\)|<|>|#|\?|\.|""|\-)", string.Empty);
        }

        // Credit: https://github.com/nikdon/SimilarityMeasure
        public static string SanitizeTitle(this string title)
        {
            // Strip all HTML
            title = Regex.Replace(title, "<[^<>]+>", "");

            // Strip numbers
            title = Regex.Replace(title, "[0-9]+", "number");

            // Strip urls
            title = Regex.Replace(title, @"(http|https)://[^\s]*", "httpaddr");

            // Strip email addresses
            title = Regex.Replace(title, @"[^\s]+@[^\s]+", "emailaddr");

            // Strip dollar sign
            title = Regex.Replace(title, "[$]+", "dollar");

            // Tokenize, get rid of any punctuation and get
            var phrases = title.Split(" @$/#.-:&*+=[]?!(){},''\">_<;%\\".ToCharArray()).Select(Sanitize).Where(x => !string.IsNullOrWhiteSpace(x));

            // Join and return
            return string.Join(" ", phrases).Trim();
        }

        private static string Sanitize(string t)
        {
            var sanitized = Regex.Replace(t, "[^a-zA-Z0-9]", string.Empty);
            var result = StopWordsList.Contains(sanitized) ? null : Stemmer.GetSteamWord(sanitized);
            return result;
        }
    }

    public class Row
    {
        public string ID_1 { get; set; }
        public string Title_1 { get; set; }
        public string ID_2 { get; set; }
        public string Title_2 { get; set; }
        public double Similarity { get; set; }
        public string Algo { get; set; }
    }
}
