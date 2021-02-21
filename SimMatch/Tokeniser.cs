namespace SimMatch
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Data;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;

    // using SemLib;

    using SimMetrics.Net.API;
    using SimMetrics.Net.Utilities;

    using StopWord;

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
