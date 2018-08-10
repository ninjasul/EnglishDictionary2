using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EnglishDictionary2
{    
    class StringUtil
    {
        public static string removeRedundantNewLineCharacters( String str )
        {
            int lastIndex = str.LastIndexOf(Environment.NewLine);
            int strLength = str.Length;

            int i = 0;
            while( Environment.NewLine.Equals(str.Substring(strLength-2-i, 2)) )
            {
                i += 2;
            }

            if( i > 0 )
            {
                str = str.Substring(0, strLength - i);
            }

            return str;
        }
    }
}
