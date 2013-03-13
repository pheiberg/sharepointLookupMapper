namespace Migrate
{
    public class MasterItemMapping
    {
        public int SourceId { get; set; }

        public string SourceTitle { get; set; }

        public int[] SourceLookupIds { get; set; }

        public int DestinationId { get; set; }

        public string DestinationTitle { get; set; }

        public int[] DestinationLookupIds { get; set; }
    }
}