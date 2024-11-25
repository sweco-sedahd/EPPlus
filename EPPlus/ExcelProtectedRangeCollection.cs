﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml
{
    public class ExcelProtectedRangeCollection : XmlHelper, IEnumerable<ExcelProtectedRange>
    {
        private readonly List<ExcelProtectedRange> _baseList = new();

        internal ExcelProtectedRangeCollection(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet ws)
            : base(nsm, topNode)
        {
            SchemaNodeOrder = ws.SchemaNodeOrder; //Fixed issue 15385
            foreach (XmlNode protectedRangeNode in topNode.SelectNodes("d:protectedRanges/d:protectedRange", nsm))
            {
                if (protectedRangeNode is not XmlElement)
                    continue;
                _baseList.Add(new ExcelProtectedRange(protectedRangeNode.Attributes["name"].Value, new ExcelAddress(SqRefUtility.FromSqRefAddress(protectedRangeNode.Attributes["sqref"].Value)), nsm, topNode));
            }
        }

        public int Count => _baseList.Count;

        public ExcelProtectedRange this[int index] => _baseList[index];

        IEnumerator<ExcelProtectedRange> IEnumerable<ExcelProtectedRange>.GetEnumerator()
        {
            return _baseList.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _baseList.GetEnumerator();
        }

        public ExcelProtectedRange Add(string name, ExcelAddress address)
        {
            if (!ExistNode("d:protectedRanges"))
            {
                CreateNode("d:protectedRanges");
            }

            foreach (ExcelProtectedRange pr in _baseList)
            {
                if (name.Equals(pr.Name, StringComparison.CurrentCultureIgnoreCase))
                {
                    throw new InvalidOperationException($"A protected range with the namn {name} already exists");
                }
            }

            XmlElement newNode = TopNode.OwnerDocument.CreateElement("protectedRange", ExcelPackage.schemaMain);
            TopNode.SelectSingleNode("d:protectedRanges", NameSpaceManager).AppendChild(newNode);
            var item = new ExcelProtectedRange(name, address, NameSpaceManager, newNode);
            _baseList.Add(item);
            return item;
        }

        public void Clear()
        {
            DeleteNode("d:protectedRanges");
            _baseList.Clear();
        }

        public bool Contains(ExcelProtectedRange item)
        {
            return _baseList.Contains(item);
        }

        public void CopyTo(ExcelProtectedRange[] array, int arrayIndex)
        {
            _baseList.CopyTo(array, arrayIndex);
        }

        public bool Remove(ExcelProtectedRange item)
        {
            DeleteAllNode("d:protectedRanges/d:protectedRange[@name='" + item.Name + "' and @sqref='" + item.Address.Address + "']");
            if (_baseList.Count == 0)
                DeleteNode("d:protectedRanges");
            return _baseList.Remove(item);
        }

        public int IndexOf(ExcelProtectedRange item)
        {
            return _baseList.IndexOf(item);
        }

        public void RemoveAt(int index)
        {
            _baseList.RemoveAt(index);
        }
    }
}