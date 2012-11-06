namespace MacysSPDL
{
    public class ColumnDefinition
    {
        public string TopicColumnName { get; set; }

        public string TopicColumnType { get; set; }

        public string LocationColumnName { get; set; }

        public string LocationColumnType { get; set; }

        public string JobCodeColumnName { get; set; }

        public string JobCodeColumnType { get; set; }

        public string SubTopicColumnName { get; set; }

        public string SubTopicColumnType { get; set; }

        public ColumnDefinition(string contentColumnName, string topicColumnName, string locationColumnName, string jobCodeColumnName, string subTopicColumnName)
        {
            TopicColumnName = GetEncodedColumnName(topicColumnName);
            LocationColumnName = GetEncodedColumnName(locationColumnName);
            JobCodeColumnName = GetEncodedColumnName(jobCodeColumnName);
            SubTopicColumnName = GetEncodedColumnName(subTopicColumnName);
            ContentColumnName = GetEncodedColumnName(contentColumnName);
        }

        private string GetEncodedColumnName(string columnName)
        {
            return columnName.Replace(" ", "_x0020_").Replace("-", "_x002d_").Replace("(", "_x0028_").Replace(")", "_x0029_");
        }

        public string ContentColumnName { get; set; }
    }
}