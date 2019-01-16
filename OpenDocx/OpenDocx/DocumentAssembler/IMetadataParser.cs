using System;
using System.Xml.Linq;

namespace OpenDocx
{
    public interface IMetadataParser
    {
        string DelimiterOpen { get; }
        string DelimiterClose { get; }
        string EmbedOpen { get; }
        string EmbedClose { get; }
        XElement TransformContentToMetadata(string content);
    }

    public class MetadataParseException : Exception
    {
        public MetadataParseException() { }
        public MetadataParseException(string message) : base(message) { }
        public MetadataParseException(string message, Exception inner) : base(message, inner) { }
    }

}
