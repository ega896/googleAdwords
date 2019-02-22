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
            var adGroups = ReadFile(_fileName).ToList();

            pb.Maximum = adGroups.Count;

            var adGroupsResult = new List<AdGroupResult>();
            ICollection<BannedCombination> bannedCombinations = new List<BannedCombination>();

            foreach (var adGroup in adGroups)
            {
                ICollection<string> wordsToRemoveFromGroup = new List<string>();

                foreach (var adWord in adGroup.AdWords)
                {
                    if (adWord.IsIgnored)
                    {
                        adGroupsResult.Add(new AdGroupResult { Name = adGroup.Name, NegativeWords = "" });
                        continue;
                    }

                    ICollection<string> wordsToRemove = new List<string>();
                    var adWordsWithoutItem = adGroups.Where(x => !x.Equals(adGroup)).SelectMany(x => x.AdWords).Where(x => !x.IsIgnored).ToList();

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
                        adWordsWithoutItem = adWordsWithoutItem.Where(x => !(x.MatchType == MatchType.PhraseMatch || x.MatchType == MatchType.BroadMatchModifier
                        && Helper.CheckPhrasesEquality(x.Value, adWord.Value))).ToList();

                        foreach (var checkItem in adWordsWithoutItem)
                        {
                            wordsToRemove.AddWords(bannedCombinations, checkItem.Value, adWord.Value);
                        }
                    }

                    foreach (var word in wordsToRemove)
                    {
                        if (wordsToRemoveFromGroup.All(x => !Helper.CheckPhrasesEquality(x, word))) wordsToRemoveFromGroup.Add(word);
                    }
                }

                adGroupsResult.Add(new AdGroupResult
                {
                    Name = adGroup.Name,
                    NegativeWords = string.Join(Environment.NewLine, wordsToRemoveFromGroup.Where(x => !adGroup.AdWords.Any(y => Helper.CheckPhrasesEquality(y.Value, x)))).TrimEnd()
                });

                pb.Increment(1);
            }

            WriteRows(_fileName, adGroupsResult);

            return "Done";
        }

        private IEnumerable<AdGroup> ReadFile(string filePath)
        {
            var existingFile = new FileInfo(filePath);

            var adGroups = new List<AdGroup>();

            using (var package = new ExcelPackage(existingFile))
            {
                var worksheet = package.Workbook.Worksheets[1];

                int dColumnRowCount = worksheet.Cells["D3:D"].Count(c => c.Value != null);
                WordsToRemove = new List<string>();
                for (int row = 3; row <= dColumnRowCount + 2; row++)
                {
                    WordsToRemove.Add(worksheet.Cells[row, 4].Value.ToString().Trim().ToLowerInvariant());
                }

                int fColumnRowCount = worksheet.Cells["F3:F"].Count(c => c.Value != null);
                Helper.HighPriorityWords = new Dictionary<int, string>();
                for (int row = 3; row <= fColumnRowCount + 2; row++)
                {
                    Helper.HighPriorityWords.Add(row - 1, worksheet.Cells[row, 6].Value.ToString().Trim().ToLowerInvariant());
                }

                int aColumnRowCount = worksheet.Cells["A3:A"].Count(c => c.Value != null);
                for (int row = 3; row <= aColumnRowCount + 2; row++)
                {
                    var adWord = new AdWord();

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

                        adWord.Value = StringWordsRemove(value.Replace("+", string.Empty)
                            .Replace("[", string.Empty)
                            .Replace("]", string.Empty)
                            .Replace("\"", string.Empty));

                        adWord.MatchType = matchType;
                        adWord.IsIgnored = adGroups.SelectMany(x => x.AdWords)
                            .Contains(adWord, new AdWordEqualityComparer());
                    }

                    var groupName = worksheet.Cells[row, 1].Value.ToString().ToLower();

                    if (adGroups.Any(x => x.Name == groupName))
                    {
                        adGroups.First(x => x.Name == groupName).AdWords.Add(adWord);
                    }
                    else
                    {
                        adGroups.Add(new AdGroup
                        {
                            Name = groupName,
                            AdWords = new List<AdWord>
                                {
                                    adWord
                                }
                        });
                    }
                }
            }

            return adGroups;
        }

        private void WriteRows(string filePath, IEnumerable<AdGroupResult> adGroupResults)
        {
            var existingFile = new FileInfo(filePath);
            using (var package = new ExcelPackage(existingFile))
            {
                var worksheet = package.Workbook.Worksheets[1];
                int rowsCount = worksheet.Cells["A3:A"].Count(c => c.Value != null) + 2;

                worksheet.Cells[1, 3].Value = "Negative words";
                foreach (var adGroupResult in adGroupResults)
                {
                    var intArray = worksheet.Cells[$"A3:A{rowsCount}"]
                        .Where(x => x.Value != null && x.Value.ToString().ToLowerInvariant() == adGroupResult.Name)
                        .Select(x => int.Parse(Regex.Match(x.Address, @"\d+").Value)).ToArray();

                    worksheet.Cells[intArray.Min(), 3, intArray.Max(), 3].Merge = true;
                    var mergedCells = worksheet.MergedCells[intArray.Max(), 3];

                    worksheet.Cells[mergedCells].Style.WrapText = true;
                    worksheet.Cells[mergedCells].Value = adGroupResult.NegativeWords;
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
