using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenDocx
{
    public interface IDataContext : IMetadataParser
    {
        string EvaluateText(string selector, bool optional);
        IDataContext[] EvaluateList(string selector, bool optional);
    }

    public class EvaluationException : Exception
    {
        public EvaluationException() { }
        public EvaluationException(string message) : base(message) { }
        public EvaluationException(string message, Exception inner) : base(message, inner) { }
    }

}
