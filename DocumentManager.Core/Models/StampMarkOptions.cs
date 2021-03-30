using System;
using System.Collections.Generic;
using System.Text;

namespace DocumentManager.Core.Models
{
    public class StampMarkOptions: DocumentOptions
    {
        public StampMarkOptions()
        {
            Text = "ＡＰＰＲＯＶＥＤ";
        }

        public StampMarkOptions(string text)
        {
            Text = string.IsNullOrEmpty(text) ? "ＡＰＰＲＯＶＥＤ" : text;
        }

        public string Text { get; set; }
    }
}
