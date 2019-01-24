using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenDocx
{
    public class AssembleResult
    {
        public WmlDocument Document { get; private set; }
        public bool HasErrors { get; private set; }

        internal AssembleResult(WmlDocument document, bool hasErrors)
        {
            Document = document;
            HasErrors = hasErrors;
        }
    }
}
