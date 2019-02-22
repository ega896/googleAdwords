using System.Collections.Generic;

namespace AdWords
{
    public class AdGroup
    {
        public string Name { get; set; }

        public ICollection<AdWord> AdWords { get; set; } = new List<AdWord>();
    }

    public class AdWord
    {
        public string Value { get; set; }

        public MatchType MatchType { get; set; }

        public bool IsIgnored { get; set; }
    }

    public class AdGroupResult
    {
        public string Name { get; set; }

        public string NegativeWords { get; set; }
    }

    public enum MatchType
    {
        BroadMatchModifier,
        PhraseMatch,
        ExactMatch
    }

    public class BannedCombination
    {
        public string CheckPhrase { get; set; }

        public string CurrentPhrase { get; set; }

        public string Combination { get; set; }
    }
}
