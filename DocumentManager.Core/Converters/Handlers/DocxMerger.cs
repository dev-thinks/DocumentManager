using Microsoft.Extensions.Logging;
using OpenXmlPowerTools;
using System.Collections.Generic;
using System.IO;

namespace DocumentManager.Core.Converters.Handlers
{
    internal class DocxMerger
    {
        private readonly ILogger _logger;

        public DocxMerger(ILogger logger)
        {
            _logger = logger;
        }

        internal void Do(string mergedTargetDoc, params string[] mergeDocs)
        {
            var sources = new List<Source>();

            foreach (var doc in mergeDocs)
            {
                if (!File.Exists(doc))
                {
                    _logger.LogError("File not exists: {FileName}", doc);
                }
                else
                {
                    sources.Add(new Source(new WmlDocument(doc), true));
                }
            }

            var mergedDoc = DocumentBuilder.BuildDocument(sources);
            mergedDoc.SaveAs(mergedTargetDoc);
        }
    }
}
