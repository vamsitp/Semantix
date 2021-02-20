namespace Semantix
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using SimMetrics.Net.Metric;

    class Program
    {
        static void Main(string[] args)
        {
            var file = args[0];
            if (File.Exists(file))
            {
                var fileService = new FileService();
                var list1 = fileService.GetRecords<dynamic>(file, true, args[1])?.ToList();
                var list2 = fileService.GetRecords<dynamic>(file, true, args[2])?.ToList();

                var titleComparisons = new List<(string Title1, string Title2, double Similarity)>();
                var s = new SmithWatermanGotoh(); // new SmithWatermanGotoh(new AffineGapRange5To0Multiplier1(), new SubCostRange5ToMinus3());
                Console.WriteLine($"list1: {list1?.Count ?? -1} | list2: {list2?.Count}");
                foreach (var t1 in list1)
                {
                    var t2 = list2.Select(l2 => new { Title = l2.Title, Similarity = s.GetSimilarity(t1.Title, l2.Title) }).OrderByDescending(l => l.Similarity).FirstOrDefault();
                    var titleComparison = $"{t1.Title} | {t2.Title} | {t2.Similarity}";
                    titleComparisons.Add((t1.Title, t2.Title, t2.Similarity));
                    Console.WriteLine(titleComparison);
                }

                fileService.WriteRecords(titleComparisons, "Comparison.xlsx");
                Console.WriteLine("Done!");
            }
        }
    }
}
