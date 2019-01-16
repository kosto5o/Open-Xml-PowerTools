/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Xml.Schema;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Packaging;
using OpenDocx;
using System.Collections;

namespace OpenDocx
{
    public class DocumentAssembler: DocumentAssemblerBase
    {
        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XmlDocument data, out bool templateError)
        {
            XDocument xDoc = data.GetXDocument();
            return AssembleDocument(templateDoc, xDoc.Root, out templateError);
        }

        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XElement data, out bool templateError)
        {
            var dataSource = new XmlDataContext(data);
            byte[] byteArray = templateDoc.DocumentByteArray;
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
                {
                    templateError = DocumentAssemblerBase.AssembleDocument(wordDoc, dataSource);
                }
                WmlDocument assembledDocument = new WmlDocument("TempFileName.docx", mem.ToArray());
                return assembledDocument;
            }
        }
    }

    public class XmlMetadataParser : IMetadataParser
    {
        public string DelimiterOpen => "<";
        public string DelimiterClose => ">";
        public string EmbedOpen => "<#";
        public string EmbedClose => "#>";

        public XElement TransformContentToMetadata(string xmlText)
        {
            XElement xml;
            try
            {
                xml = XElement.Parse(xmlText);
            }
            catch (XmlException e)
            {
                throw new MetadataParseException(e.Message, e.InnerException);
            }
            return xml;
        }
    }

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

        public string EvaluateText(string xPath, bool optional )
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
                var selectedData = ((IEnumerable) xPathSelectResult).Cast<XObject>();
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
                
                if (selectedDatum is XElement) return ((XElement) selectedDatum).Value;

                if (selectedDatum is XAttribute) return ((XAttribute) selectedDatum).Value;
            }

            return xPathSelectResult.ToString();

        }

        public void Release()
        {
            _element = null;
        }
    }
}
