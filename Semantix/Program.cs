namespace Semantix
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text;

    using ColoredConsole;

    using SemLib;

    using SimMetrics.Net.API;

    class Program
    {

        private const string Space = " ";
        private static readonly KeywordAnalyzer ka = new KeywordAnalyzer();
        private static readonly FileService fileService = new FileService();

        private readonly static string[] Exclusions = new[] { "ChapmanLengthDeviation", "Jaro", "JaroWinkler", "MongeElkan" }; // , "NeedlemanWunch"
        private static Dictionary<string, IStringMetric> Algos = Assembly
            .GetExecutingAssembly().GetReferencedAssemblies()
            .Select(x => Assembly.Load(x)).SelectMany(x => x.GetTypes()
            .Where(t => !t.IsAbstract && t.IsAssignableTo(typeof(IStringMetric)) && !Exclusions.Any(t.Name.Equals)))
            .ToDictionary(x => x.Name, x => (IStringMetric)Activator.CreateInstance(x));

        static void Main(string[] args)
        {
            foreach (var algo in Algos)
            {
                Console.WriteLine($"{algo.Key} - {algo.Value.ToString()}");
            }

            var file = string.Empty;
            Dictionary<string, string> list1;
            Dictionary<string, string> list2;

            if (args.Length <= 0)
            {
                ColorConsole.Write("path-to-excel-file (containing the 2 sheets to compare - with ID & Title columns) <space> max-matches per row: ");
                file = Console.ReadLine();
            }
            else
            {
                file = args[0];
            }

            if (File.Exists(file))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                var output = $"{Path.GetFileNameWithoutExtension(file)}_Comparison.xlsx";

                var excelData = fileService.GetRecords(file);

                var table1 = excelData.Tables[0];
                var table2 = excelData.Tables[1];

                if (table1.Columns.Count > 1)
                {
                    list1 = table1.Rows.Cast<DataRow>()?.ToDictionary(r => r[0]?.ToString(), r => r[1]?.ToString()); // fileService.GetRecords<dynamic>(file, true, args[1])?.ToList();
                    list2 = table2.Rows.Cast<DataRow>()?.ToDictionary(r => r[0]?.ToString(), r => r[1]?.ToString()); // fileService.GetRecords<dynamic>(file, true, args[2])?.ToList();
                }
                else if (table1.Columns.Count == 1)
                {
                    list1 = table1.Rows.Cast<DataRow>()?.Select((r, i) => (r, i)).ToDictionary(row => row.i.ToString(), row => row.r[0]?.ToString()); // fileService.GetRecords<dynamic>(file, true, args[1])?.ToList();
                    list2 = table2.Rows.Cast<DataRow>()?.Select((r, i) => (r, i)).ToDictionary(row => row.i.ToString(), row => row.r[0]?.ToString()); // fileService.GetRecords<dynamic>(file, true, args[2])?.ToList();
                }
                else
                {
                    ColorConsole.WriteLine("Input Excel file should have 2 sheets to compare with ID & Title columns!");
                    return;
                }

                var max = 1;
                if (args.Length > 1 && !int.TryParse(args[1], out max))
                {
                    max = 1;
                }

                var titleComparisons = new List<Row>();
                ColorConsole.WriteLine($"L1: {list1?.Count ?? -1}".White(), " | ", $"L2: {list2?.Count}".Blue(), Environment.NewLine);
                for (var i = 0; i < list1.Count; i++)
                {
                    var t1 = list1.ElementAt(i);
                    var matches = list2.Select(l2 =>
                    {
                        var row = new Row { ID_1 = t1.Key, Title_1 = t1.Value, ID_2 = l2.Key, Title_2 = l2.Value };
                        SetSimilarity(row);
                        return row;
                    }).OrderByDescending(l => l.Similarity).Take(max);

                    foreach (var match in matches)
                    {
                        titleComparisons.Add(match);
                        ColorConsole.WriteLine(string.Empty.PadLeft(5), $"{i}. L1: {match.ID_1} - {match.Title_1}".White());
                        ColorConsole.WriteLine(string.Empty.PadLeft(5), $"{i}. L2: {match.ID_2} - {match.Title_2}".Blue());
                        ColorConsole.WriteLine(string.Empty.PadLeft(5), $"{i}. SM: {match.Similarity} ({match.Algo})".Color(match.Similarity > 0.5 ? ConsoleColor.Green : ConsoleColor.Red));
                        ColorConsole.WriteLine();
                    }
                }

                fileService.WriteRecords(titleComparisons, output);
                ColorConsole.WriteLine("Done!");
                Process.Start(new ProcessStartInfo(output) { UseShellExecute = true });
            }
        }

        private static void SetSimilarity(Row row)
        {
            var t1a = string.Join(Space, ka.Analyze(row.Title_1).Keywords.Where(k => !k.Word.Contains(Space)).Select(k => k.Word)); // k.Rank > 0.0m
            var t2a = string.Join(Space, ka.Analyze(row.Title_2).Keywords.Where(k => !k.Word.Contains(Space)).Select(k => k.Word)); // k.Rank > 0.0m

            // ColorConsole.WriteLine(t1a.DarkGray(), " | ", t2a.DarkGray());
            var similarity = Algos.Select(a => (Similarity: a.Value.GetSimilarity(t1a, t2a), Algo: a.Key)).OrderByDescending(s => s.Similarity).FirstOrDefault();
            row.Similarity = similarity.Similarity;
            row.Algo = similarity.Algo;
        }
    }
}
