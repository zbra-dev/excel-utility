using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace ExcelUtility.Utils
{
    public class XElementData
    {
        private XElement data;

        public string DefaultNamespace { get; private set; }
        public string Value { get { return data.Value; } set { data.Value = value; } }
        public string this[string attributeName] { get { return AttributeValue(attributeName); } set { SetAttributeValue(attributeName, value); } }

        private XElementData(XElement data, string defaultNamespace)
        {
            DefaultNamespace = defaultNamespace;
            this.data = data;
        }

        public XElementData(string prefix, XElement data)
        {
            this.data = data;
            DefaultNamespace = GetPrefixNamespace(prefix);
        }

        public XElementData(XElement data)
        {
            this.data = data;
            DefaultNamespace = data.GetDefaultNamespace().ToString();
        }

        private XElementData(string name, string defaultNamespace)
        {
            DefaultNamespace = defaultNamespace;
            data = new XElement(XName.Get(name, DefaultNamespace));
        }

        private string GetPrefixNamespace(string prefix)
        {
            var prefixNamespace = data.GetNamespaceOfPrefix(prefix);
            if (prefixNamespace == null)
                throw new ArgumentException(string.Format("Couldn't find prefix {0}", prefix), "prefix");
            return prefixNamespace.ToString();
        }

        private XElementData New(string prefix, string name)
        {
            return new XElementData(name: name, defaultNamespace: GetPrefixNamespace(prefix));
        }

        private XElementData New(string name)
        {
            return new XElementData(name: name, defaultNamespace: DefaultNamespace);
        }

        public IEnumerable<XElementData> Descendants(string name)
        {
            return data.Descendants(XName.Get(name, DefaultNamespace)).Select(d => new XElementData(d, DefaultNamespace));
        }

        public XElementData Element(string prefix, string name)
        {
            var prefixNamespace = GetPrefixNamespace(prefix);
            var elementData = data.Element(XName.Get(name, prefixNamespace));
            return elementData == null ? null : new XElementData(elementData, prefixNamespace);
        }

        public XElementData Element(string name)
        {
            var elementData = data.Element(XName.Get(name, DefaultNamespace));
            return elementData == null ? null : new XElementData(elementData, DefaultNamespace);
        }

        public XElementData ElementAt(string path)
        {
            if (string.IsNullOrEmpty(path))
                return null;
            return ElementAt(path.Split('.'), 0);
        }

        private XElementData ElementAt(string[] path, int index)
        {
            if (index >= path.Length)
                return null;
            var split = path[index].Split(':');
            var next = split.Length == 2 ? Element(split[0], split[1]) : Element(path[index]);
            if (next == null || index + 1 == path.Length)
                return next;
            return next.ElementAt(path, index + 1);
        }

        public string AttributeValue(string prefix, string name)
        {
            var attribute = data.Attribute(XName.Get(name, GetPrefixNamespace(prefix)));
            return attribute == null ? null : attribute.Value;
        }

        public string AttributeValue(string name)
        {
            var attribute = data.Attribute(XName.Get(name));
            return attribute == null ? null : attribute.Value;
        }

        public void Save(string filePath)
        {
            data.Save(filePath);
        }

        public void RemoveAttribute(string name)
        {
            data.SetAttributeValue(XName.Get(name), null);
        }

        public void RemoveAttribute(string prefix, string name)
        {
            data.SetAttributeValue(XName.Get(name, GetPrefixNamespace(prefix)), null);
        }

        public void SetAttributeValue(string prefix, string name, object value)
        {
            data.SetAttributeValue(XName.Get(name, GetPrefixNamespace(prefix)), value);
        }

        public void SetAttributeValue(string name, object value)
        {
            data.SetAttributeValue(XName.Get(name), value);
        }

        public void SetAttributeValues(string values)
        {
            foreach (var value in values.Split(' ').Where(s => !string.IsNullOrEmpty(s)))
            {
                var split = value.Split('=');
                if (split.Length == 2)
                    SetAttributeValue(split[0], split[1]);
            }
        }

        public void SetElementValue(string name, object value)
        {
            data.SetElementValue(XName.Get(name, DefaultNamespace), value);
        }

        public XElementData Add(string prefix, string name)
        {
            var content = New(prefix, name);
            Add(content);
            return content;
        }

        public XElementData Add(string name)
        {
            var content = New(name);
            Add(content);
            return content;
        }

        private void Add(XElementData content)
        {
            data.Add(content.data);
        }

        public XElementData AddAfterSelf(string name)
        {
            var newData = New(name);
            AddAfterSelf(newData);
            return newData;
        }

        private void AddAfterSelf(XElementData content)
        {
            data.AddAfterSelf(content.data);
        }

        public XElementData AddBeforeSelf(string name)
        {
            var newData = New(name);
            AddBeforeSelf(newData);
            return newData;
        }

        private void AddBeforeSelf(XElementData content)
        {
            data.AddBeforeSelf(content.data);
        }

        public void RemoveNodes()
        {
            data.RemoveNodes();
        }

        public void Remove()
        {
            data.Remove();
        }

        public override string ToString()
        {
            return data.ToString();
        }
    }
}
