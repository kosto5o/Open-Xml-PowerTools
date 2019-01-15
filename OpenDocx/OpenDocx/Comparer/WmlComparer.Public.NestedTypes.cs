﻿using System;
using System.Xml.Linq;

namespace OpenDocx
{
    public static partial class WmlComparer
    {
        public class WmlComparerRevision
        {
            public WmlComparerRevisionType RevisionType;
            public string Text;
            public string Author;
            public string Date;
            public XElement ContentXElement;
            public XElement RevisionXElement;
            public Uri PartUri;
            public string PartContentType;
        }

        public enum WmlComparerRevisionType
        {
            Inserted,
            Deleted
        }
    }
}
