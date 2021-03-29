using System.Collections.Generic;

namespace DocumentManager.Core.Models
{
    public class Placeholders: DocumentOptions
    {
        /// <summary>
        /// .ctor
        /// </summary>
        public Placeholders()
        {
            NewLineTag = "<br/>";
            TextPlaceholders = new Dictionary<string, string>();
            TablePlaceholders = new List<TableElement>();
            ImagePlaceholders = new Dictionary<string, ImageElement>();
            HyperlinkPlaceholders = new Dictionary<string, HyperlinkElement>();
            IsWaterMarkNeeded = false;
        }

        /// <summary>
        /// NewLineTags are important only for .docx as input. If you use .html as input, then just use "<br/>"
        /// </summary>
        public string NewLineTag { get; set; }

        public Dictionary<string, string> TextPlaceholders { get; set; }

        public List<TableElement> TablePlaceholders { get; set; }

        public Dictionary<string, ImageElement> ImagePlaceholders { get; set; }

        public Dictionary<string, HyperlinkElement> HyperlinkPlaceholders { get; set; }

        public bool IsWaterMarkNeeded { get; set; }
    }
}
