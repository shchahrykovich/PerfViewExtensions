using Reporting.Implementations;
using Reporting.Viewers;

namespace Reporting.Tests.Viewers
{
    internal class FakeViewer : IViewer
    {
        public Statistics Stat { get; set; }

        public void Show(Statistics stat)
        {
            Stat = stat;
        }
    }
}
