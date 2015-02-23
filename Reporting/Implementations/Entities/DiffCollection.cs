using System.Collections.Generic;

namespace Reporting.Implementations.Entities
{
    internal class DiffCollection
    {
        public List<Diff> Diffs { get; set; }

        public DiffCollection()
        {
            Diffs = new List<Diff>();
        }
    }
}
