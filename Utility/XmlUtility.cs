using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace AUTOSYS.Utility
{
    public static class XmlUtility
    {
        /// <summary>
        /// Net-KINDサーバーのIP取得
        /// </summary>
        /// <param name="code"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static string GetXmValue(string key)
            => XElement.Load(AppDomain.CurrentDomain.BaseDirectory + @"\setting.xml")
            .Element(key).Attribute("value").Value;

        /// <summary>
        /// Net-KINDサーバーのIP取得
        /// </summary>
        /// <param name="code"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static void SetXmValue(string key, string value)
        {
            XElement ex = XElement.Load(AppDomain.CurrentDomain.BaseDirectory + @"\setting.xml");
            ex.Element(key).SetAttributeValue("value", value ?? "");
            ex.Save(AppDomain.CurrentDomain.BaseDirectory + @"\setting.xml");
        }

        /// <summary>
        /// Net-KINDサーバーのIP取得
        /// </summary>
        /// <param name="code"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static void SetXmValue(Dictionary<string, string> dic)
        {
            XElement ex = XElement.Load(AppDomain.CurrentDomain.BaseDirectory + @"\setting.xml");
            foreach (string key in dic.Keys)
            {
                ex.Element(key).SetAttributeValue("value", dic[key] ?? "");
            }
            ex.Save(AppDomain.CurrentDomain.BaseDirectory + @"\setting.xml");
        }

        /// <summary>
        /// Net-KINDサーバーのIP取得
        /// </summary>
        /// <param name="code"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static Dictionary<string, string> GetXmDataValues(string key)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            XElement ex = XElement.Load(AppDomain.CurrentDomain.BaseDirectory + @"\setting.xml").Element(key);
            if (ex != null)
                ex.Elements("data").ToList().ForEach(e => result.Add(e.Attribute("name").Value, e.Attribute("value").Value));

            return result;
        }

        /// <summary>
        /// Net-KINDサーバーのIP取得
        /// </summary>
        /// <param name="code"></param>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static void SetXmlDataValues(string key, Dictionary<string, string> dic)
        {
            XElement orig = XElement.Load(AppDomain.CurrentDomain.BaseDirectory + @"\setting.xml");
            XElement ex = orig.Element(key);

            if (ex == null)
            {
                XElement nd = new XElement("data");
                foreach (string k in dic.Keys)
                {
                    nd.SetAttributeValue(k, dic[k] ?? "");
                }
                orig.Add(nd);
            }
            else
            {
                foreach (string k in dic.Keys)
                {
                    XElement e = ex.Elements("data").Where(x => x.Attribute("name").Value == k).FirstOrDefault();
                    if (e != null)
                    {
                        e.SetAttributeValue("value", dic[k] ?? "");
                    }
                    else
                    {
                        XElement node = new XElement(key);
                        node.Attribute("name").Value = k;
                        node.Attribute("value").Value = dic[k];
                        ex.Add(node);
                    }
                }

                orig.Save(AppDomain.CurrentDomain.BaseDirectory + @"\setting.xml");
            }
        }

        /// <summary>
        /// メッセージ取得
        /// </summary>
        /// <returns></returns>
        public static List<XElement> GetXElements(string path, string group)
            => XElement.Load(path).Elements(group).ToList();

        public static List<XElement> GetXElements(XElement parent, string group)
            => parent.Elements(group).ToList();

        public static XElement GetXElement(string path, string group)
            => XElement.Load(path).Element(group);
    }
}
