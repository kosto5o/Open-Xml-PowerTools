using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Threading.Tasks;

namespace OpenDocx
{
    public class XmlDataContext : XmlMetadataParser, IDataContext
    {
        private XElement _element;

        public XmlDataContext(XElement data)
        {
            _element = data;
        }

        public IDataContext[] EvaluateList(string selector)
        {
            IEnumerable<XElement> repeatingData;
            try
            {
                repeatingData = _element.XPathSelectElements(selector);
            }
            catch (XPathException e)
            {
                throw new EvaluationException("XPathException: " + e.Message);
            }
            var newContent = repeatingData.Select(d =>
            {
                return new XmlDataContext(d);
            })
                .ToArray();
            return newContent;
        }

        public string EvaluateText(string xPath, bool optional)
        {
            object xPathSelectResult;
            try
            {
                //support some cells in the table may not have an xpath expression.
                if (String.IsNullOrWhiteSpace(xPath)) return String.Empty;

                xPathSelectResult = _element.XPathEvaluate(xPath);
            }
            catch (XPathException e)
            {
                throw new EvaluationException("XPathException: " + e.Message, e);
            }

            if ((xPathSelectResult is IEnumerable) && !(xPathSelectResult is string))
            {
                var selectedData = ((IEnumerable)xPathSelectResult).Cast<XObject>();
                if (!selectedData.Any())
                {
                    if (optional) return string.Empty;
                    throw new EvaluationException(string.Format("XPath expression ({0}) returned no results", xPath));
                }
                if (selectedData.Count() > 1)
                {
                    throw new EvaluationException(string.Format("XPath expression ({0}) returned more than one node", xPath));
                }

                XObject selectedDatum = selectedData.First();

                if (selectedDatum is XElement) return ((XElement)selectedDatum).Value;

                if (selectedDatum is XAttribute) return ((XAttribute)selectedDatum).Value;
            }

            return xPathSelectResult.ToString();

        }

        public void Release()
        {
            _element = null;
        }
    }

    public class AsyncXmlDataContext : XmlMetadataParser, IAsyncDataContext
    {
        private XElement _element;

        public AsyncXmlDataContext(XElement data)
        {
            _element = data;
        }

        public async Task<IAsyncDataContext[]> EvaluateListAsync(string selector)
        {
            await Task.Yield(); // make this async -- silly, but really just for testing
            IEnumerable<XElement> repeatingData;
            try
            {
                repeatingData = _element.XPathSelectElements(selector);
            }
            catch (XPathException e)
            {
                throw new EvaluationException("XPathException: " + e.Message);
            }
            var newContent = repeatingData.Select(d =>
                {
                    return new AsyncXmlDataContext(d);
                })
                .ToArray();
            return newContent;
        }

        public async Task<string> EvaluateTextAsync(string xPath, bool optional)
        {
            await Task.Yield(); // make this async -- silly, but really just for testing
            object xPathSelectResult;
            try
            {
                //support some cells in the table may not have an xpath expression.
                if (String.IsNullOrWhiteSpace(xPath)) return String.Empty;

                xPathSelectResult = _element.XPathEvaluate(xPath);
            }
            catch (XPathException e)
            {
                throw new EvaluationException("XPathException: " + e.Message, e);
            }

            if ((xPathSelectResult is IEnumerable) && !(xPathSelectResult is string))
            {
                var selectedData = ((IEnumerable)xPathSelectResult).Cast<XObject>();
                if (!selectedData.Any())
                {
                    if (optional) return string.Empty;
                    throw new EvaluationException(string.Format("XPath expression ({0}) returned no results", xPath));
                }
                if (selectedData.Count() > 1)
                {
                    throw new EvaluationException(string.Format("XPath expression ({0}) returned more than one node", xPath));
                }

                XObject selectedDatum = selectedData.First();

                if (selectedDatum is XElement) return ((XElement)selectedDatum).Value;

                if (selectedDatum is XAttribute) return ((XAttribute)selectedDatum).Value;
            }

            return xPathSelectResult.ToString();

        }

        public async Task ReleaseAsync()
        {
            await Task.Yield(); // make this async -- silly, but really just for testing
            _element = null;
        }
    }
}
