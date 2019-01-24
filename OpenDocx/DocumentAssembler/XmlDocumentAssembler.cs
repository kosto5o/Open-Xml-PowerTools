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
using System.Xml;
using System.Xml.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;

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

        public class AsyncAssembleResult
        {
            public WmlDocument assembledDocument;
            public bool templateError;
        }

        public static Task<AsyncAssembleResult> AssembleDocumentAsync(WmlDocument templateDoc, XmlDocument data)
        {
            XDocument xDoc = data.GetXDocument();
            return AssembleDocumentAsync(templateDoc, xDoc.Root);
        }

        public static async Task<AsyncAssembleResult> AssembleDocumentAsync(WmlDocument templateDoc, XElement data)
        {
            System.Diagnostics.Debug.WriteLine(templateDoc.FileName);
            var dataSource = new AsyncXmlDataContext(data);
            byte[] byteArray = templateDoc.DocumentByteArray;
            var result = new AsyncAssembleResult();
            using (MemoryStream mem = new MemoryStream())
            {
                mem.Write(byteArray, 0, (int)byteArray.Length);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
                {
                    result.templateError = await DocumentAssemblerBase.AssembleDocumentAsync(wordDoc, dataSource);
                }
                result.assembledDocument = new WmlDocument("TempFileName.docx", mem.ToArray());
                return result;
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
}
