namespace Semantix
{
    using System.Collections.Generic;
    using System.Data;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;

    using CsvHelper;
    using CsvHelper.Configuration;
    using CsvHelper.Excel;

    using ExcelDataReader;

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
