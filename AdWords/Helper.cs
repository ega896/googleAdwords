using System.Collections.Generic;
using System.Linq;

namespace AdWords
{
    public static class Helper
    {
        public static void AddWords(this ICollection<string> words, ICollection<BannedCombination> bannedCombinations,
            string currentPhrase, string phraseToCheck)
        {
            var phraseToCheckWords = phraseToCheck.Split(' ');
            var currentPhraseWords = currentPhrase.Split(' ');

            var exceptionPhrase = string.Join(" ", currentPhraseWords.Except(phraseToCheckWords));
            var intersectionPhrase = (exceptionPhrase + " " + phraseToCheck).Trim();

            if (!new BannedCombination { CurrentPhrase = currentPhrase, CheckPhrase = phraseToCheck, Combination = intersectionPhrase }.IsInBannedCombinationCollection(bannedCombinations) &&
                !new BannedCombination{ CurrentPhrase = phraseToCheck, CheckPhrase = currentPhrase, Combination = intersectionPhrase}.IsInBannedCombinationCollection(bannedCombinations) &&
                currentPhraseWords.Length >= phraseToCheckWords.Length && GetPriorityKey(currentPhrase) == GetPriorityKey(phraseToCheck) &&
                currentPhrase != phraseToCheck)
            {
                if (!CheckPhrasesEquality(currentPhrase, intersectionPhrase))
                {
                    bannedCombinations.Add(new BannedCombination
                    {
                        CheckPhrase = phraseToCheck,
                        CurrentPhrase = currentPhrase,
                        Combination = intersectionPhrase
                    });
                }
            }

            if (!intersectionPhrase.IsInCollection(words) &&
                !new BannedCombination{ CurrentPhrase = phraseToCheck, CheckPhrase = currentPhrase, Combination = intersectionPhrase}.IsInBannedCombinationCollection(bannedCombinations) &&
                phraseToCheckWords.Length <= currentPhraseWords.Length)
            {

                bool addWord = true;

                if (phraseToCheckWords.Length == currentPhraseWords.Length)
                {
                    addWord = !(currentPhrase != phraseToCheck && GetPriorityKey(phraseToCheck) < GetPriorityKey(currentPhrase));
                }
                    
                if (intersectionPhrase != phraseToCheck && addWord) words.Add(intersectionPhrase);
            }
        }

        public static bool IsInBannedCombinationCollection(this BannedCombination bc, ICollection<BannedCombination> bcs)
        {
            return bcs.Any(x => CheckPhrasesEquality(x.Combination, bc.Combination) && x.CheckPhrase == bc.CheckPhrase && x.CurrentPhrase == bc.CurrentPhrase);
        }

        public static bool IsInCollection(this string phrase, IEnumerable<string> stringsCol)
        {
            var phraseWords = phrase.Split(' ');
            var stringWords = stringsCol.Select(x => x.Split(' '));

            return stringWords.Any(stringWord => stringWord.Intersect(phraseWords).Count() == stringWord.Length &&
                                                 stringWord.Intersect(phraseWords).Count() == phraseWords.Length);
        }

        public static bool CheckPhrasesEquality(string phrase1, string phrase2)
        {
            var phrase1Words = phrase1.Split(' ');
            var phrase2Words = phrase2.Split(' ');

            return phrase1Words.Intersect(phrase2Words).Count() == phrase1Words.Length &&
                   phrase2Words.Intersect(phrase1Words).Count() == phrase2Words.Length;
        }

        public static int GetPriorityKey(string phrase)
        {
            return
                HighPriorityWords.Any(stringWord => phrase.Contains(stringWord.Value)) ?
                    HighPriorityWords.Where(stringWord => phrase.Contains(stringWord.Value)).Min(x => x.Key) : 10000;
        }

        public static Dictionary<int, string> HighPriorityWords = new Dictionary<int, string>();
    }
}
