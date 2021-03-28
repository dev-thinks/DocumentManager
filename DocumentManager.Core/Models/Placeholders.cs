﻿using System.Collections.Generic;

namespace DocumentManager.Core.Models
{
    public class Placeholders
    {
        /// <summary>
        /// .ctor
        /// </summary>
        public Placeholders()
        {
            NewLineTag = "<br/>";
            TextPlaceholderStartTag = "##";
            TextPlaceholderEndTag = "##";
            TablePlaceholderStartTag = "==";
            TablePlaceholderEndTag = "==";
            ImagePlaceholderStartTag = "++";
            ImagePlaceholderEndTag = "++";
            HyperlinkPlaceholderStartTag = "//";
            HyperlinkPlaceholderEndTag = "//";

            TextPlaceholders = new Dictionary<string, string>();
            TablePlaceholders = new List<TableElement>();
            ImagePlaceholders = new Dictionary<string, ImageElement>();
            HyperlinkPlaceholders = new Dictionary<string, HyperlinkElement>();
        }

        /// <summary>
        /// NewLineTags are important only for .docx as input. If you use .html as input, then just use "<br/>"
        /// </summary>
        public string NewLineTag { get; set; }

        /// <summary>
        /// Start and End Tags can e. g. be both "##"
        /// A placeholder could be ##TextPlaceHolder##
        /// </summary>
        public string TextPlaceholderStartTag { get; set; }

        public string TextPlaceholderEndTag { get; set; }

        public Dictionary<string, string> TextPlaceholders { get; set; }

        /// <summary>
        /// For tables it works that way: * 1. If you have a table in the word document, create 1 row with a different Dictionary keys
        /// * Then e.g.you want to have 10 rows in the end, you add 10 values to each array of the Dictionary value
        /// A placeholder could be ==TextPlaceHolder==
        /// Start and End Tags can e. g. be both "=="
        /// </summary>
        public string TablePlaceholderStartTag { get; set; }

        public string TablePlaceholderEndTag { get; set; }

        public List<TableElement> TablePlaceholders { get; set; }

        /// <summary>
        /// Important: The MemoryStream may carry an image. Allowed file types: JPEG/JPG, BMP, TIFF, GIF, PNG
        /// Take different replacement tags here, else there may be collision with the text replacements, e. g. "++" 
        /// </summary>
        public string ImagePlaceholderStartTag { get; set; }

        public string ImagePlaceholderEndTag { get; set; }

        public Dictionary<string, ImageElement> ImagePlaceholders { get; set; }


        public string HyperlinkPlaceholderStartTag { get; set; }

        public string HyperlinkPlaceholderEndTag { get; set; }

        public Dictionary<string, HyperlinkElement> HyperlinkPlaceholders { get; set; }
    }
}
