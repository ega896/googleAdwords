using System.Collections.Generic;

namespace AdWords
{
    public class AdWordEqualityComparer : IEqualityComparer<AdWord>
    {
        public bool Equals(AdWord x, AdWord y)
        {
            return x.MatchType == y.MatchType && Helper.CheckPhrasesEquality(x.Value, y.Value);
        }

        public int GetHashCode(AdWord obj)
        {
            return obj.MatchType.GetHashCode() * 17 + obj.Value.GetHashCode();
        }
    }
}
