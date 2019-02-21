using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace Console
{
    public class Program
    {
        public static void Main(string[] args)
        {
            System.Console.WriteLine("Enter filepath: ");
            var filePath = System.Console.ReadLine();
            var adWords = ReadFile(filePath).ToList();

            var resultCollection = new List<string>();
            ICollection<BannedCombination> bannedCombinations = new List<BannedCombination>();

            foreach (var adWord in adWords)
            {
                if (adWord.IsIgnored)
                {
                    resultCollection.Add("");
                    continue;
                }

                ICollection<string> wordsToRemove = new List<string>();

                var adWordsWithoutItem = adWords.Where(x => !x.Equals(adWord) && !x.IsIgnored).ToList();

                if (adWord.MatchType == MatchType.BroadMatchModifier)
                {
                    foreach (var checkItem in adWordsWithoutItem)
                    {
                        wordsToRemove.AddWords(bannedCombinations, checkItem.Value, adWord.Value);

                        if (checkItem.Value != adWord.Value || wordsToRemove.Contains(checkItem.Value)) continue;

                        switch (checkItem.MatchType)
                        {
                            case MatchType.ExactMatch:
                                wordsToRemove.Add($"[{checkItem.Value}]");
                                break;
                            case MatchType.PhraseMatch:
                                wordsToRemove.Add($"\"{checkItem.Value}\"");
                                break;
                        }
                    }
                }
                else if (adWord.MatchType == MatchType.PhraseMatch)
                {
                    foreach (var checkItem in adWordsWithoutItem)
                    {
                        wordsToRemove.AddWords(bannedCombinations, checkItem.Value, adWord.Value);

                        if (checkItem.MatchType == MatchType.ExactMatch && checkItem.Value == adWord.Value &&
                            !wordsToRemove.Contains(checkItem.Value))
                        {
                            wordsToRemove.Add($"[{checkItem.Value}]");
                        }
                    }
                }
                else if (adWord.MatchType == MatchType.ExactMatch)
                {
                    adWordsWithoutItem = adWordsWithoutItem.Where(x => !(x.MatchType == MatchType.PhraseMatch || x.MatchType == MatchType.BroadMatchModifier && Helper.CheckPhrasesEquality(x.Value, adWord.Value))).ToList();

                    foreach (var checkItem in adWordsWithoutItem)
                    {
                        wordsToRemove.AddWords(bannedCombinations, checkItem.Value, adWord.Value);
                    }
                }

                resultCollection.Add(string.Join(Environment.NewLine, wordsToRemove).TrimEnd());
            }

            WriteRows(filePath, resultCollection);

            System.Console.WriteLine("done");
            System.Console.ReadKey();
        }

        public static IEnumerable<AdWord> ReadFile(string filePath)
        {
            var existingFile = new FileInfo(filePath);
            var adWords = new List<AdWord>();

            using (var package = new ExcelPackage(existingFile))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.End.Row;

                for (int row = 3; row <= rowCount; row++)
                {
                    if (worksheet.Cells[row, 2].Value != null)
                    {
                        var value = worksheet.Cells[row, 2].Value.ToString().Trim().ToLowerInvariant();
                        MatchType matchType;
                        if (value.Contains('\"'))
                        {
                            matchType = MatchType.PhraseMatch;
                        }
                        else if (value.Contains('['))
                        {
                            matchType = MatchType.ExactMatch;
                        }
                        else
                        {
                            matchType = MatchType.BroadMatchModifier;
                        }

                        var adWord = new AdWord
                        {
                            Value = StringWordsRemove(value.Replace("+", string.Empty)
                                .Replace("[", string.Empty)
                                .Replace("]", string.Empty)
                                .Replace("\"", string.Empty)),
                            MatchType = matchType
                        };
                        adWord.IsIgnored = adWords.Contains(adWord, new AdWordEqualityComparer());

                        adWords.Add(adWord);
                    }
                    if (worksheet.Cells[row, 4].Value != null)
                    {
                        WordsToRemove.Add(worksheet.Cells[row, 4].Value.ToString().Trim().ToLowerInvariant());
                    }
                    if (worksheet.Cells[row, 6].Value != null)
                    {
                        Helper.HighPriorityWords.Add(row-1, worksheet.Cells[row, 6].Value.ToString().Trim().ToLowerInvariant());
                    }
                }
            }

            return adWords;
        }

        public static IDictionary<string, string> WriteRows(string filePath, ICollection<string> collection)
        {
            var existingFile = new FileInfo(filePath);
            var result = new Dictionary<string, string>();

            using (var package = new ExcelPackage(existingFile))
            {
                var worksheet = package.Workbook.Worksheets[0];
                worksheet.Cells[1, 3].Value = "Negative words";

                for (int i = 3; i < collection.Count + 3; i++)
                {
                    worksheet.Cells[i, 3].Style.WrapText = true;
                    worksheet.Cells[i, 3].Value = collection.ElementAtOrDefault(i-3);                 
                }

                package.Save();
            }

            return result;
        }

        public static ICollection<string> WordsToRemove = new List<string>();

        public static string StringWordsRemove(string stringToClean)
        {
            var regex = new Regex("[ ]{2,}", RegexOptions.None);
            var value = regex.Replace(string.Join(" ", stringToClean.Split(' ').Except(WordsToRemove)), " ");
            return value;
        }
    }
}
