namespace SimMatch
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Data;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using System.Text.RegularExpressions;

    using ColoredConsole;

    // using SemLib;

    using SimMetrics.Net.API;
    using SimMetrics.Net.Utilities;

    using StopWord;

    class Program
    {

        private const string Space = " ";
        private const double DefaultMinThreshold = 0.5;
        private const int DefaultMaxSimilarities = 1;
        // private static readonly KeywordAnalyzer ka = new KeywordAnalyzer();
        private static readonly FileService fileService = new FileService();
        private static readonly SimMatchTokeniser Tokeniser = new SimMatchTokeniser();

        private readonly static List<string> Exclusions = new List<string>(); // { "ChapmanLengthDeviation", "Jaro", "JaroWinkler", "MongeElkan" }; // , "NeedlemanWunch"
        private static Dictionary<string, IStringMetric> Algos = Assembly
            .GetExecutingAssembly().GetReferencedAssemblies()
            .Select(x => Assembly.Load(x)).SelectMany(x => x.GetTypes()
            .Where(t => !t.IsAbstract && t.IsAssignableTo(typeof(IStringMetric)))) // && !Exclusions.Any(t.Name.Equals)
            .ToDictionary(x => x.Name, x =>
            {
                try
	            {
                    var instance = (IStringMetric)Activator.CreateInstance(x, Tokeniser);
                    // ColorConsole.WriteLine($"{x.Name} created with Tokeniser");
                    return instance;
                }
	            catch
	            {
                    var instance = (IStringMetric)Activator.CreateInstance(x);
                    Exclusions.Add(x.Name);
                    // ColorConsole.WriteLine($"{x.Name} created without Tokeniser".DarkYellow());
                    return instance;
                }
            });
        private static List<Process> outputProcs = new List<Process>();

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
            }
            catch (Exception ex)
            {
                ColorConsole.WriteLine(ex.Message.White().OnRed(), Environment.NewLine);
            }
        }

        private static void Process(Dictionary<string, string> list1, Dictionary<string, string> list2, IEnumerable<(KeyValuePair<string, IStringMetric> Algo, int Index)> allAlgos, string file, double threshold, int max)
        {
            ColorConsole.WriteLine();
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
                    var rows = algos.Select(a =>
                    {
                        var row = new Row { ID_1 = t1.Key, Title_1 = t1.Value, ID_2 = l2.Key, Title_2 = l2.Value };
                        SetSimilarity(row, a);
                        return row;
                    }).OrderByDescending(l => l.Similarity);

                    return rows;
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

            var output = $"{Path.GetFileNameWithoutExtension(file)}_Comparison_{string.Join("_", filtered.Select(x => x.Index)).Trim('_')}.xlsx";
            var outputProc = outputProcs.SingleOrDefault(o => o?.StartInfo?.FileName?.Contains(output) == true);
            if (outputProc != null)
            {
                outputProc.Kill(); // outputProc?.Close();
                outputProc.WaitForExit();
                outputProcs.Remove(outputProc);
            }

            fileService.WriteRecords(titleComparisons, output);
            ColorConsole.WriteLine(" --------------------------------------------------");
            outputProcs.Add(System.Diagnostics.Process.Start(new ProcessStartInfo(output) { UseShellExecute = true }));
        }

        private static void SetSimilarity(Row row, KeyValuePair<string, IStringMetric> algo)
        {
            //var t1a = string.Join(Space, ka.Analyze(row.Title_1).Keywords.Where(k => !k.Word.Contains(Space)).Select(k => k.Word)); // k.Rank > 0.0m
            //var t2a = string.Join(Space, ka.Analyze(row.Title_2).Keywords.Where(k => !k.Word.Contains(Space)).Select(k => k.Word)); // k.Rank > 0.0m

            //// ColorConsole.WriteLine(t1a.DarkGray(), " | ", t2a.DarkGray());
            //var similarity = algo.Value.GetSimilarity(t1a, t2a);
            var similarity = algo.Value.GetSimilarity(row.Title_1, row.Title_2);
            row.Similarity = similarity;
            row.Algo = algo.Key;
        }
    }

    public sealed class SimMatchTokeniser : ITokeniser
    {
        private readonly TokeniserUtilities<string> tokenUtilities = new TokeniserUtilities<string>();

        // Credit: https://github.com/nikdon/SimilarityMeasure
        public Collection<string> Tokenize(string word)
        {
            var wordCountList = new Dictionary<string, int>();
            var collection = new Collection<string>();
            var sanitized = Sanitize(word);
            foreach (string part in sanitized)
            {
                // Strip non-alphanumeric characters
                string stripped = Regex.Replace(part, "[^a-zA-Z0-9]", "");
                if (!StopWordHandler.IsWord(stripped.ToLower()))
                {
                    try
                    {
                        var stem = new Annytab.Stemmer.EnglishStemmer().GetSteamWord(stripped);
                        if (stem.Length > 0)
                        {
                            // Build the word count list
                            if (wordCountList.ContainsKey(stem))
                            {
                                wordCountList[stem]++;
                            }
                            else
                            {
                                wordCountList.Add(stem, 0);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Tokenizer exception source: {0}", e.Message);
                    }
                }
            }

            var vocabList = wordCountList.Where(w => w.Value >= 0); // threshold
            foreach (var item in vocabList)
            {
                collection.Add(item.Key);
            }

            return collection;
        }

        private static string[] Sanitize(string text)
        {
            // Strip all HTML
            text = Regex.Replace(text, "<[^<>]+>", "");

            // Strip numbers
            text = Regex.Replace(text, "[0-9]+", "number");

            // Strip urls
            text = Regex.Replace(text, @"(http|https)://[^\s]*", "httpaddr");

            // Strip email addresses
            text = Regex.Replace(text, @"[^\s]+@[^\s]+", "emailaddr");

            // Strip dollar sign
            text = Regex.Replace(text, "[$]+", "dollar");

            // Tokenize and also get rid of any punctuation
            return text.Split(" @$/#.-:&*+=[]?!(){},''\">_<;%\\".ToCharArray());
        }

        public Collection<string> TokenizeToSet(string word)
        {
            if (word != null)
            {
                return tokenUtilities.CreateSet(Tokenize(word));
            }
            return null;
        }

        public string Delimiters { get; } = "\r\n\t \x00a0";

        public string ShortDescriptionString => "SimMatchTokeniser";

        public ITermHandler StopWordHandler { get; set; } = new SimMatchStopTermHandler();
    }

    public sealed class SimMatchStopTermHandler : ITermHandler
    {
        private static readonly List<string> StopWordsList = StopWords.GetStopWords("en").ToList();

        public void AddWord(string termToAdd)
        {
            StopWordsList.Add(termToAdd);
        }

        public bool IsWord(string termToTest)
        {
            return StopWordsList.Contains(termToTest.ToLower());
        }

        public void RemoveWord(string termToRemove)
        {
            StopWordsList.Remove(termToRemove);
        }

        public int NumberOfWords => 0;

        public string ShortDescriptionString => "SimMatchStopTermHandler";

        public StringBuilder WordsAsBuffer => new StringBuilder();
    }
}
