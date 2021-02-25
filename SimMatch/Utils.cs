namespace SimMatch
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text.RegularExpressions;

    using Annytab.Stemmer;

    using ClosedXML.Excel;

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

        public void WriteRecords(IEnumerable<Row> records, string outputPath)
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            GenerateExcel(records, outputPath);
            //using (var excelWriter = new ExcelWriter(outputPath, CultureInfo.InvariantCulture))
            //{
            //    excelWriter.WriteRecords(records);
            //}
        }

        private static void GenerateExcel(IEnumerable<Row> records, string outputPath)
        {
            var wb = new XLWorkbook { ReferenceStyle = XLReferenceStyle.Default, CalculateMode = XLCalculateMode.Auto };
            var ws = wb.Worksheets.Add("SimMatch");
            ws.Style.Font.FontName = "Segoe UI";
            ws.Style.Font.FontSize = 10.00;

            var table = ws.Cell(1, 1).InsertTable(records.ToDataTable(), "SimMatch");
            ws.Row(1).Style.Font.Bold = true;

            var titleGroups = records.GroupBy(x => x.Title_1);
            foreach (var group in titleGroups)
            {
                var maxCountTitle2 = group.GroupBy(s => s.Title_2).OrderByDescending(s => s.Count()).First().Key;
                var title1Rows = table.Rows().Where(r => r.Cell(2).Value.Equals(group.Key) && r.Cell(4).Value.Equals(maxCountTitle2));
                foreach (var row in title1Rows)
                {
                    row.Style.Font.Bold = true;
                }
            }

            table.Theme = XLTableTheme.TableStyleLight13;
            // ws.Columns().AdjustToContents();

            wb.SaveAs(outputPath);
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

        // Credit: https://stackoverflow.com/questions/18100783/how-to-convert-a-list-into-data-table
        public static DataTable ToDataTable<T>(this IEnumerable<T> records)
        {
            var properties = TypeDescriptor.GetProperties(typeof(T));
            var table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
            {
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            }

            foreach (T item in records)
            {
                var row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                {
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                }

                table.Rows.Add(row);
            }

            return table;
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
