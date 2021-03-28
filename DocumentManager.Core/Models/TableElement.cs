using System.Collections.Generic;

namespace DocumentManager.Core.Models
{
    public class TableElement
    {
        public TableElement()
        {
            RowValues = new Dictionary<string, string[]>();
            TableName = null;
            Prefix = "TableStart";
            Suffix = "TableEnd";
        }

        public string TableName { get; set; }

        public string Prefix { get; set; }

        public string Suffix { get; set; }

        public Dictionary<string, string[]> RowValues { get; set; }
    }
}
