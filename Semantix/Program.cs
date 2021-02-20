namespace Semantix
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;

    using ColoredConsole;

    using SemLib;

    using SimMetrics.Net.API;
    using SimMetrics.Net.Metric;
    using SimMetrics.Net.Utilities;

    class Program
    {
        private const string Space = " ";
        private static readonly KeywordAnalyzer ka = new KeywordAnalyzer();
        private static readonly FileService fileService = new FileService();
        private static readonly Dictionary<string, AbstractStringMetric> Algos = new Dictionary<string, AbstractStringMetric>
        {
            { nameof(SmithWatermanGotoh), new SmithWatermanGotoh() }, // new SmithWatermanGotoh(new AffineGapRange5To0Multiplier1(), new SubCostRange5ToMinus3());
            { nameof(Levenstein), new Levenstein() },
            { nameof(SmithWaterman), new SmithWaterman() },
            { nameof(MatchingCoefficient), new MatchingCoefficient() },
            { nameof(OverlapCoefficient), new OverlapCoefficient() },
        };

        static void Main(string[] args)
        {
            var file = args[0];
            if (File.Exists(file))
            {
                var output = $"{Path.GetFileNameWithoutExtension(file)}_Comparison.xlsx";

                var list1 = fileService.GetRecords<dynamic>(file, true, args[1])?.ToList();
                var list2 = fileService.GetRecords<dynamic>(file, true, args[2])?.ToList();
                var max = 1;

                if (args.Length > 3 && !int.TryParse(args[3], out max))
                {
                    max = 1;
                }

                var titleComparisons = new List<Row>();
                ColorConsole.WriteLine($"L1: {list1?.Count ?? -1}".White(), " | ", $"L2: {list2?.Count}".Blue(), Environment.NewLine);
                foreach (var l1 in list1.Select((x, i) => (t1: x, i: i + 1)))
                {
                    var i = l1.i;
                    var t1 = l1.t1;
                    var matches = list2.Select(l2 =>
                    {
                        var row = new Row { ID_1 = t1.ID, Title_1 = t1.Title, ID_2 = l2.ID, Title_2 = l2.Title };
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
