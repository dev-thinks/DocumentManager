using System;

namespace DocumentManager.Core.Models
{
    public class DocumentOptions
    {
        public DocumentOptions()
        {
            var currentGuid = Guid.NewGuid();

            WorkingLocation = "";
            OpenOfficeLocation = "";
            CanCleanUpMarkup = false;
            CurrentProcessId = currentGuid;
        }

        public string WorkingLocation { get; set; }

        public string OpenOfficeLocation { get; set; }

        public bool CanCleanUpMarkup { get; set; }

        public Guid CurrentProcessId { get; set; }
    }
}
