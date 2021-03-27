using System.IO;

namespace DocumentManager.Core.Models
{
    public class ImageElement
    {
        public MemoryStream MemStream { get; set; }

        public double Dpi { get; set; } // Dots per inch
    }
}
