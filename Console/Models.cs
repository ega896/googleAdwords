namespace Console
{
    public class AdWord
    {
        public string Value { get; set; }

        public MatchType MatchType { get; set; }

        public bool IsIgnored { get; set; }
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
