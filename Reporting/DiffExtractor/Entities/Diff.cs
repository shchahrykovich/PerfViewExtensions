namespace Reporting.DiffExtractor.Entities
{
    internal class Diff
    {
        public readonly DataRecord Start;
        public readonly DataRecord Finish;

        public Diff(DataRecord start, DataRecord finish)
        {
            Start = start;
            Finish = finish;
        }
    }
}