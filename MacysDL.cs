using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using MacysSPDL.ListsWebSvc;

namespace MacysSPDL
{
    public class MacysDL
    {
        private readonly string _listName;
        private readonly Lists _listsSvc;
        private readonly ColumnDefinition _columnDefinition;

        //find replace xmlnamespace..
        private static readonly Regex StripXmlnsRegex = new Regex(@"(xmlns:?[^=]*=[""][^""]*[""])", RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Compiled | RegexOptions.CultureInvariant);

        //to find and replace short XmlNameSpace like z:,rs: from xmlresponse by GetListItems method
        private static readonly Regex RemoveShortNameSpace = new Regex(@"(?<lessthan><)(?<closetag>[/])?(?<shortname>\w+:)\s*", RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public MacysDL(string spSiteURL, string listName, NetworkCredential credential, ColumnDefinition columnDefinition)
        {
            _listsSvc = new Lists
                            {
                                Url = spSiteURL + "/_vti_bin/lists.asmx",
                                Credentials = credential,
                                AllowAutoRedirect = true
                            };
            _listName = listName;
            _columnDefinition = columnDefinition;
            if (String.IsNullOrEmpty(_columnDefinition.JobCodeColumnType))
                _columnDefinition.JobCodeColumnType = "Text";
            if (String.IsNullOrEmpty(_columnDefinition.LocationColumnType))
                _columnDefinition.LocationColumnType = "Text";
            if (String.IsNullOrEmpty(_columnDefinition.TopicColumnType))
                _columnDefinition.TopicColumnType = "Text";
            if (String.IsNullOrEmpty(_columnDefinition.SubTopicColumnType))
                _columnDefinition.SubTopicColumnType = "Text";
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="topic"></param>
        /// <param name="jobCode"></param>
        /// <param name="location"></param>
        /// <returns></returns>
        public string GetContents(string topic, string jobCode, string location)
        {
            var xmlDoc = new XmlDocument();
            XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");
            XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");
            XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");
            ndQueryOptions.InnerXml = "<IncludeMandatoryColumns>TRUE</IncludeMandatoryColumns>" + "<DateInUtc>TRUE</DateInUtc>";
            ndViewFields.InnerXml = @"<FieldRef Name='" + _columnDefinition.ContentColumnName + @"' />";
            ndQuery.InnerXml = GetQuery(topic, location, jobCode);
            XmlNode ndListItems = _listsSvc.GetListItems(_listName, null, ndQuery, ndViewFields, null, ndQueryOptions, null);

            //remove namespace xmlnls from  xml..
            string xmlResponse = StripXmlnsRegex.Replace(ndListItems.InnerXml, "");

            //find and replace short XmlNameSpace like z:,rs: from responce with space..
            xmlResponse = RemoveShortNameSpace.Replace(xmlResponse, delegate(Match m)
            {
                Group closetag = m.Groups["closetag"];
                if (closetag.Length != 0)
                    return "</";
                return "<";
            });

            //load xml from removed XmlNameSpace and short name of XmlNameSpace..
            var resultxmlDoc = XDocument.Parse(xmlResponse);

            //iterate each row in sharepoint list.
            //in result xml each row is in element "row"
            var items = from item in resultxmlDoc.XPathSelectElements("//row")
                        let attribute = item.Attribute("ows_" + _columnDefinition.ContentColumnName)
                        where attribute != null
                        select new
                        {
                            //get Title Field of SharePoint list...
                            Content = Convert.ToString(attribute.Value)
                        };
            string outPutContent = string.Empty;

            //display each item in title field in console..
            Array.ForEach(items.ToArray(), item => outPutContent += item.Content);

            return outPutContent;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="jobCode"></param>
        /// <param name="location"></param>
        /// <returns></returns>
        public List<string> GetAllContents(string jobCode, string location)
        {
            var contents = new List<string>();
            var xmlDoc = new XmlDocument();
            XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");
            XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");
            XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");
            ndQueryOptions.InnerXml = "<IncludeMandatoryColumns>TRUE</IncludeMandatoryColumns>" + "<DateInUtc>TRUE</DateInUtc>";

            ndViewFields.InnerXml = @"<FieldRef Name='" + _columnDefinition.ContentColumnName + @"' />";

            ndQuery.InnerXml = GetQuery(location, jobCode);
            XmlNode ndListItems = _listsSvc.GetListItems(_listName, null, ndQuery, ndViewFields, null, ndQueryOptions, null);

            //remove namespace xmlnls from  xml..
            string xmlResponse = StripXmlnsRegex.Replace(ndListItems.InnerXml, "");

            //find and replace short XmlNameSpace like z:,rs: from responce with space..
            xmlResponse = RemoveShortNameSpace.Replace(xmlResponse, delegate(Match m)
            {
                Group closetag = m.Groups["closetag"];
                if (closetag.Length != 0)
                    return "</";
                return "<";
            });

            //load xml from removed XmlNameSpace and short name of XmlNameSpace..
            var resultxmlDoc = XDocument.Parse(xmlResponse);

            //iterate each row in sharepoint list.
            //in result xml each row is in element "row"
            var items = from item in resultxmlDoc.XPathSelectElements("//row")
                        let attribute = item.Attribute("ows_" + _columnDefinition.ContentColumnName)
                        where attribute != null
                        select new
                        {
                            //get Title Field of SharePoint list...
                            Content = Convert.ToString(attribute.Value)
                        };

            //display each item in title field in console..
            Array.ForEach(items.ToArray(), item => contents.Add(item.Content));
            return contents;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="jobCode"></param>
        /// <param name="location"></param>
        /// <returns></returns>
        public List<string> GetTopicsAndSubTopics(string jobCode, string location)
        {
            var xmlDoc = new XmlDocument();
            XmlNode ndQuery = xmlDoc.CreateNode(XmlNodeType.Element, "Query", "");
            XmlNode ndViewFields = xmlDoc.CreateNode(XmlNodeType.Element, "ViewFields", "");
            XmlNode ndQueryOptions = xmlDoc.CreateNode(XmlNodeType.Element, "QueryOptions", "");
            ndQueryOptions.InnerXml = string.Empty;
            ndViewFields.InnerXml = @"
                                        <FieldRef Name='" + _columnDefinition.TopicColumnName + @"' />
                                        <FieldRef Name='" + _columnDefinition.SubTopicColumnName + @"' />
                                      ";
            ndQuery.InnerXml = GetQuery(location, jobCode);
            XmlNode ndListItems = _listsSvc.GetListItems(_listName, null, ndQuery, ndViewFields, null, ndQueryOptions, null);

            //remove namespace xmlnls from  xml..
            string xmlResponse = StripXmlnsRegex.Replace(ndListItems.InnerXml, "");

            //find and replace short XmlNameSpace like z:,rs: from responce with space..
            xmlResponse = RemoveShortNameSpace.Replace(xmlResponse, delegate(Match m)
            {
                Group closetag = m.Groups["closetag"];

                if (closetag.Length != 0)
                    return "</";
                return "<";
            });

            //load xml from removed XmlNameSpace and short name of XmlNameSpace..
            var resultxmlDoc = XDocument.Parse(xmlResponse);

            //iterate each row in sharepoint list.
            //in result xml each row is in element "row"
            return (from row in resultxmlDoc.XPathSelectElements("//row") where row.Attribute("ows_" + _columnDefinition.TopicColumnName) != null && row.Attribute("ows_" + _columnDefinition.SubTopicColumnName) != null let topic = row.Attribute("ows_" + _columnDefinition.TopicColumnName) let subTopic = row.Attribute("ows_" + _columnDefinition.SubTopicColumnName) where topic != null && subTopic != null select topic.Value.Substring(topic.Value.IndexOf("#", StringComparison.Ordinal) + 1) + ":" + subTopic.Value.Substring(subTopic.Value.IndexOf("#", StringComparison.Ordinal) + 1)).ToList();
        }

        internal string GetQuery(string topic, string location, string jobCode)
        {
            // Return list item collection based on the document name
            var stringBuilder = new StringBuilder();
            stringBuilder.Append(@"

                                    <Where>
                                        <And>
                                            <And>
                                                <Eq>
                                                    <FieldRef Name='" + _columnDefinition.TopicColumnName + "'/><Value Type='" + _columnDefinition.TopicColumnType + "'>" + topic + @"</Value>
                                                </Eq>
                                                <Eq>
                                                    <FieldRef Name='" + _columnDefinition.LocationColumnName + "'/><Value Type='" + _columnDefinition.LocationColumnType + "'>" + location + @"</Value>
                                                </Eq>
                                            </And>
                                            <Eq>
                                                <FieldRef Name='" + _columnDefinition.JobCodeColumnName + "'/><Value Type='" + _columnDefinition.JobCodeColumnType + "'>" + jobCode + @"</Value>
                                            </Eq>
                                        </And>
                                    </Where>
                            ");
            return stringBuilder.ToString();
        }

        internal string GetQuery(string location, string jobCode)
        {
            // Return list item collection based on the document name
            var stringBuilder = new StringBuilder();
            stringBuilder.Append(@"

                            <Where>
                                <And>
                                    <Eq>
                                        <FieldRef Name='" + _columnDefinition.LocationColumnName + "'/><Value Type='" + _columnDefinition.LocationColumnType + "'>" + location + @"</Value>
                                    </Eq>
                                    <Eq>
                                        <FieldRef Name='" + _columnDefinition.JobCodeColumnName + "'/><Value Type='" + _columnDefinition.JobCodeColumnType + "'>" + jobCode + @"</Value>
                                    </Eq>
                                </And>
                            </Where>
                        ");
            return stringBuilder.ToString();
        }
    }
}