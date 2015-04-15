using Reporting.Implementations;

namespace Reporting.Viewers
{
    internal interface IViewer
    {
        void Show(Statistics stat);
    }
}
