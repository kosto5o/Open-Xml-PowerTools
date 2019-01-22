using System;
using System.Threading.Tasks;

namespace OpenDocx
{
    public interface IDataContext : IMetadataParser
    {
        string EvaluateText(string selector, bool optional);
        IDataContext[] EvaluateList(string selector);
        void Release();
    }

    public interface IAsyncDataContext : IMetadataParser
    {
        Task<string> EvaluateTextAsync(string selector, bool optional);
        Task<IAsyncDataContext[]> EvaluateListAsync(string selector);
        Task ReleaseAsync();
    }

    public class EvaluationException : Exception
    {
        public EvaluationException() { }
        public EvaluationException(string message) : base(message) { }
        public EvaluationException(string message, Exception inner) : base(message, inner) { }
    }

}
