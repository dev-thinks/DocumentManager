using System;

namespace DocumentManager.Core.Models
{
    public class DocumentOptions
    {
        public DocumentOptions()
        {
            var currentGuid = Guid.NewGuid();

            TempLocation = $"working\\{currentGuid}";
            OpenOfficeLocation = "";
            CanCleanUpMarkup = false;
            CurrentProcessId = currentGuid;
        }

        public string TempLocation { get; set; }

        public string OpenOfficeLocation { get; set; }

        public bool CanCleanUpMarkup { get; set; }

        public Guid CurrentProcessId { get; set; }
    }
}
