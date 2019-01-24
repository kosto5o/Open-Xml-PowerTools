using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenDocx
{
    public class CompileResult
    {
        public WmlDocument CompiledTemplate { get; private set; }
        public bool HasErrors { get; private set; }

        internal CompileResult(WmlDocument compiledTemplate, bool hasErrors)
        {
            CompiledTemplate = compiledTemplate;
            HasErrors = hasErrors;
        }
    }
}
