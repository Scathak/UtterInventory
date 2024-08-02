using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Collections.Generic;
using System.Xml.Serialization;
using System.IO;
using System.Xml;

namespace UtterInventory
{
    public partial class ThisAddIn
    {
        public class TableXml
        {
            [XmlElement(ElementName = "name")]
            public string TableName { get; set; }

            [XmlElement(ElementName = "column")]
            public string[] ColumnsNames { get; set; }
        }
        [XmlRoot(ElementName = "structure")]
        public class Structure
        {
            [XmlIgnore]
            [XmlAttribute(AttributeName = "xmlns:xsd")]
            public string Xsd { get; set; }

            [XmlIgnore]
            [XmlAttribute(AttributeName = "xmlns:xsi")]
            public string Xsi { get; set; }

            [XmlElement(ElementName = "table")]
            public List<TableXml> Tables { get; set; }
        }
        public string ToXml(Structure source)
        {
            using (StringWriter writer = new StringWriter())
            {
                XmlSerializer serializer = new XmlSerializer(source.GetType());
                serializer.Serialize(writer, source);
                return writer.ToString();
            }
        }
        public static object FromXml<T>(string source)
        {
            object obj;
            using (var textReader = new StringReader(source))
            {
                using (XmlTextReader reader = new XmlTextReader(textReader))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(T));
                    obj = (T)serializer.Deserialize(reader);
                }
            }
            return obj;
        }
        public Structure ReadCustomXML(Workbook wb, string context)
        {
            Structure returnObject = null;
            CustomXMLParts customElements = wb.CustomXMLParts;

            foreach (CustomXMLPart element in customElements)
            {
                string xml = element.XML;
                if (xml.Contains("<?xml version") && xml.Contains(context))
                {
                    returnObject = (Structure)FromXml<Structure>(xml);
                }
            }
            return returnObject;
        }
        public void defaultStructureToCustomXml(Workbook wb)
        {
            var ListOfTables = new Structure();
            ListOfTables.Tables = new List<TableXml>();
            foreach (var item in Globals.ThisAddIn.TablesStructure)
            {
                var table = new TableXml();
                table.TableName = item.Key;
                table.ColumnsNames = item.Value;
                ListOfTables.Tables.Add(table);
            }
            wb.CustomXMLParts.Add(ToXml(ListOfTables));
        }
    }
}
