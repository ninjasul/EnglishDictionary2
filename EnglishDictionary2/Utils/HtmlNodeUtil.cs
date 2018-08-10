using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EnglishDictionary2
{
    class HtmlNodeUtil
    {
        public static bool isNullOrEmpty(HtmlNode htmlNode)
        {
            return (htmlNode == null || string.IsNullOrEmpty(htmlNode.InnerText));
        }
    }
}
