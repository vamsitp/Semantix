namespace SimMatch
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

    // using SemLib;

    using SimMetrics.Net.API;

    class Program
    {
        private const double DefaultMinThreshold = 0.5;
        private const int DefaultMaxSimilarities = 1;

        // private const string Space = " ";
        // private static readonly KeywordAnalyzer ka = new KeywordAnalyzer();
        // private static readonly SimMatchTokeniser Tokeniser = new SimMatchTokeniser();

        private static readonly FileService fileService = new FileService();
        private static readonly List<Process> OutputProcs = new List<Process>();
        private readonly static List<string> Exclusions = new List<string>(); // { "ChapmanLengthDeviation", "Jaro", "JaroWinkler", "MongeElkan" }; // , "NeedlemanWunch"
        private static readonly Dictionary<string, IStringMetric> Algos = Assembly
            .GetExecutingAssembly().GetReferencedAssemblies()
            .Select(x => Assembly.Load(x)).SelectMany(x => x.GetTypes()
            .Where(t => !t.IsAbstract && t.IsAssignableTo(typeof(IStringMetric)))) // && !Exclusions.Any(t.Name.Equals)
            .ToDictionary(x => x.Name, x =>
            {
             //   try
	         //   {
             //       var instance = (IStringMetric)Activator.CreateInstance(x, Tokeniser);
             //       // ColorConsole.WriteLine($"{x.Name} created with Tokeniser");
             //       return instance;
             //   }
	         //   catch
	         //   {
                    var instance = (IStringMetric)Activator.CreateInstance(x);
                    // Exclusions.Add(x.Name);
                    // ColorConsole.WriteLine($"{x.Name} created without Tokeniser".DarkYellow());
                    return instance;
             //   }
            });

        static void Main(string[] input)
        {
            try
            {
                var args = input?.ToList();
                var file = string.Empty;
                Dictionary<string, string> list1;
                Dictionary<string, string> list2;

                if (args.Count <= 0)
                {
                    ColorConsole.Write("path-to-excel-file (containing the 2 sheets to compare - with ID & Title columns): ".Cyan());
                    args.Add(Console.ReadLine().Trim(new[] { ' ', '"' }));
                    ColorConsole.Write("min-threshold b/w 0.0 - 1.0 (default is 0.5): ".Cyan());
                    args.Add(Console.ReadLine().Trim(new[] { ' ', '"' }));
                    ColorConsole.Write("max-similarity-matches in the 2nd list per each row in the 1st list (default is 1): ".Cyan());
                    args.Add(Console.ReadLine()?.Trim(new[] { ' ', '"' }));
                }

                file = args[0];
                if (File.Exists(file))
                {
                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

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

                    var threshold = DefaultMinThreshold;
                    if (args.Count > 1 && !double.TryParse(args[1], out threshold))
                    {
                        threshold = DefaultMinThreshold;
                    }

                    var max = DefaultMaxSimilarities;
                    if (args.Count > 2 && !int.TryParse(args[2], out max))
                    {
                        max = DefaultMaxSimilarities;
                    }

                    var allAlgos = Algos.Select((x, i) => (Algo: x, Index: i + 1));
                    while (true)
                    {
                        try
                        {
                            Process(list1, list2, allAlgos, file, threshold, max);
                        }
                        catch (Exception ex)
                        {
                            ColorConsole.WriteLine(ex.Message.Red(), Environment.NewLine);
                        }
                    }
                }
                else
                {
                    ColorConsole.WriteLine($"Invalid file: {file}".Red());
                }
            }
            catch (Exception ex)
            {
                ColorConsole.WriteLine(ex.Message.White().OnRed(), Environment.NewLine);
            }
        }

        private static void Process(Dictionary<string, string> list1, Dictionary<string, string> list2, IEnumerable<(KeyValuePair<string, IStringMetric> Algo, int Index)> allAlgos, string file, double threshold, int max)
        {
            ColorConsole.WriteLine();
            ColorConsole.WriteLine($"0. ".PadLeft(5).DarkCyan(), "G_Diff_Match_Patch".DarkGray());
            foreach (var item in allAlgos)
            {
                ColorConsole.WriteLine($"{item.Index}. ".PadLeft(5).Cyan(), item.Algo.Key.Color(Exclusions.Any(item.Algo.Key.Equals) ? ConsoleColor.DarkGray : ConsoleColor.White));
            }

            ColorConsole.WriteLine("\n Enter the index of the Algos to run (space or comma separated)".Cyan());
            ColorConsole.Write(" > ".Cyan());
            var inputs = Console.ReadLine()?.Split(new[] { ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries)?.Select(x => x.Trim());
            var filtered = (inputs?.Count() > 0 ? allAlgos.Where(x => inputs.Any(x.Index.ToString().Equals)) : allAlgos.Where(x => !Exclusions.Any(x.Algo.Key.Equals)));
            var algos = filtered.Select(x => x.Algo);

            var titleComparisons = new List<Row>();
            ColorConsole.WriteLine($"\n L1: {list1?.Count ?? -1}".White(), " | ", $"L2: {list2?.Count}".Blue(), Environment.NewLine, " --------------------------------------------------".Cyan());
            for (var i = 0; i < list1.Count; i++)
            {
                var t1 = list1.ElementAt(i);
                var matches = list2.SelectMany(l2 =>
                {
                    if (inputs?.Count() > 0 && !inputs.Any(x => x.Equals("0")))
                    {
                        var rows = algos.Select(a =>
                        {
                            var row = new Row { ID_1 = t1.Key, Title_1 = t1.Value, ID_2 = l2.Key, Title_2 = l2.Value };
                            SetSimilarity(row, a);
                            return row;
                        }).OrderByDescending(l => l.Similarity);

                        return rows;
                    }
                    else
                    {
                        return GetDmpMatches(t1, l2);
                    }
                }).Where(m => m.Similarity >= threshold).OrderByDescending(l => l.Similarity).Take(max);

                foreach (var match in matches)
                {
                    titleComparisons.Add(match);
                    ColorConsole.WriteLine($" {i}. L1: {match.ID_1} - {match.Title_1}".White());
                    ColorConsole.WriteLine($" {i}. L2: {match.ID_2} - {match.Title_2}".Blue());
                    ColorConsole.WriteLine($" {i}. SM: {match.Similarity} ({match.Algo})".Color(match.Similarity >= threshold ? ConsoleColor.Green : ConsoleColor.Red));
                    ColorConsole.WriteLine();
                }
            }

            var output = $"{Path.GetFileNameWithoutExtension(file)}_Comparison_{string.Join("_", inputs).Trim('_')}.xlsx";
            var outputProc = OutputProcs.SingleOrDefault(o => o?.StartInfo?.FileName?.Contains(output) == true);
            if (outputProc != null)
            {
                outputProc.Kill(); // outputProc?.Close();
                outputProc.WaitForExit();
                OutputProcs.Remove(outputProc);
            }

            fileService.WriteRecords(titleComparisons, output);
            ColorConsole.WriteLine(" --------------------------------------------------");
            OutputProcs.Add(System.Diagnostics.Process.Start(new ProcessStartInfo(output) { UseShellExecute = true }));
        }

        private static void SetSimilarity(Row row, KeyValuePair<string, IStringMetric> algo)
        {
            //var t1a = string.Join(Space, ka.Analyze(row.Title_1).Keywords.Where(k => !k.Word.Contains(Space)).Select(k => k.Word)); // k.Rank > 0.0m
            //var t2a = string.Join(Space, ka.Analyze(row.Title_2).Keywords.Where(k => !k.Word.Contains(Space)).Select(k => k.Word)); // k.Rank > 0.0m

            //// ColorConsole.WriteLine(t1a.DarkGray(), " | ", t2a.DarkGray());
            //var similarity = algo.Value.GetSimilarity(t1a, t2a);

            var similarity = algo.Value.GetSimilarity(row.Title_1.SanitizeTitle(), row.Title_2.SanitizeTitle());
            row.Similarity = similarity;
            row.Algo = algo.Key;
        }

        private static IEnumerable<Row> GetDmpMatches(KeyValuePair<string, string> t1, KeyValuePair<string, string> l2)
        {
            var rows = new List<Row>();
            var match = GetDmpMatch(t1.Value, l2.Value);
            if (match > 1)
            {
                var row = new Row { ID_1 = t1.Key, Title_1 = t1.Value, ID_2 = l2.Key, Title_2 = l2.Value, Similarity = match, Algo = nameof(DiffMatchPatch) };
                rows.Add(row);
            }

            return rows.OrderByDescending(l => l.Similarity);
        }

        // https://github.com/google/diff-match-patch/wiki/Language:-C%23
        private static int GetDmpMatch(string title1, string title2)
        {
            var dmp = new DiffMatchPatch.diff_match_patch();
            var diff = dmp.diff_main(title1.SanitizeTitle(), title2.SanitizeTitle()); // var diff = dmp.match_main(title2.SanitizeTitle(), title1.SanitizeTitle(), 0);
            dmp.diff_cleanupSemantic(diff);
            var equals = diff.Count(x => x.text.Length > 3 && x.operation == DiffMatchPatch.Operation.EQUAL); // 3 variable
            if (Debugger.IsAttached && equals > 0)
            {
                Console.WriteLine(string.Join(Environment.NewLine, diff.Where(x => x.operation == DiffMatchPatch.Operation.EQUAL).Select(d => $" {d.operation} - {d.text}")).Trim());
            }

            return equals;
        }
    }
}
