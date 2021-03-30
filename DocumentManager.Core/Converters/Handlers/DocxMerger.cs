using Microsoft.Extensions.Logging;
using OpenXmlPowerTools;
using System.Collections.Generic;

namespace DocumentManager.Core.Converters.Handlers
{
    internal class DocxMerger
    {
        public DocxMerger(ILogger logger)
        {

        }

        internal void Do(string mergedTargetDoc, params string[] mergeDocs)
        {
            var sources = new List<Source>();

            foreach (var doc in mergeDocs)
            {
                sources.Add(new Source(new WmlDocument(doc), true));
            }

            var mergedDoc = DocumentBuilder.BuildDocument(sources);
            mergedDoc.SaveAs(mergedTargetDoc);
        }
    }
}
