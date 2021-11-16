using System;
using System.Collections.Generic;

namespace SearcherLibrary
{
    public interface IFileSearchHandler
    {

        List<MatchedLine> Search(String fileName, IEnumerable<String> searchTerms, Matcher matcher);

    }
}
