using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OfficeOpenXml;

namespace AdWords
{
    public class ExcelService
    {
        private readonly string _fileName;

        public static ICollection<string> WordsToRemove = new List<string>();

        public ExcelService(string fileName)
        {
            _fileName = fileName;
        }

        public string Execute(ProgressBar pb)
        {
            var adWords = ReadFile(_fileName).ToList();

            pb.Maximum = adWords.Count;

            var resultCollection = new List<string>();
            ICollection<BannedCombination> bannedCombinations = new List<BannedCombination>();

            foreach (var adWord in adWords)
            {
                if (adWord.IsIgnored)
                {
                    resultCollection.Add("");
                    pb.Increment(1);
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
                pb.Increment(1);
            }

            WriteRows(_fileName, resultCollection);

            return "Done";
        }

        private IEnumerable<AdWord> ReadFile(string filePath)
        {
            var existingFile = new FileInfo(filePath);
            var adWords = new List<AdWord>();

            using (var package = new ExcelPackage(existingFile))
            {
                var worksheet = package.Workbook.Worksheets[1];
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
                        Helper.HighPriorityWords.Add(row - 1, worksheet.Cells[row, 6].Value.ToString().Trim().ToLowerInvariant());
                    }
                }
            }

            return adWords;
        }

        private void WriteRows(string filePath, ICollection<string> collection)
        {
            var existingFile = new FileInfo(filePath);
            using (var package = new ExcelPackage(existingFile))
            {
                var worksheet = package.Workbook.Worksheets[1];
                worksheet.Cells[1, 3].Value = "Negative words";

                for (int i = 3; i < collection.Count + 3; i++)
                {
                    worksheet.Cells[i, 3].Style.WrapText = true;
                    worksheet.Cells[i, 3].Value = collection.ElementAtOrDefault(i - 3);
                }

                package.Save();
            }
        }

        public static string StringWordsRemove(string stringToClean)
        {
            var regex = new Regex("[ ]{2,}", RegexOptions.None);
            var value = regex.Replace(string.Join(" ", stringToClean.Split(' ').Except(WordsToRemove)), " ");
            return value;
        }
    }
}
