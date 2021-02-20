namespace Semantix
{
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;

    using CsvHelper;
    using CsvHelper.Configuration;
    using CsvHelper.Excel;

    public class FileService
    {
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
            using (var excelWriter = new ExcelWriter(outputPath, CultureInfo.InvariantCulture))
            {
                excelWriter.WriteRecords(records);
            }
        }
    }

    public static class SemantixExtensions
    {
        public static string SanitizeHeader(this string header)
        {
            return Regex.Replace(header, @"(\s+|@|&|'|\(|\)|<|>|#|\?|\.|""|\-)", string.Empty);
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
