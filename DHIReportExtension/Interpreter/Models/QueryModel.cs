namespace DHIWordExtension.Interpreter.Models
{
    public class QueryModel
    {
        public QueryType QueryType { get; set; }
        public string DataBaseName { get; set; }
        public string TableName { get; set; }
        public string Filter { get; set; }
        public string Select { get; set; }
        public string Sort { get; set; }
    }
}