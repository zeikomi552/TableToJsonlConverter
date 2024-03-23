using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableToJsonlConverter.Conveters
{
    public class ZkJsonlBase
    {
        public Dictionary<string, string> EscapeStringsDic { get; private set; }
            = new Dictionary<string, string>()
            {
                {"\\", "\\\\" },
                {"\"", "\\\""},
                {"/", "\\/" },
                {"\b", "\\b" },
                {"\f", "\\f" },
                {"\n", "\\n" },
                {"\r", "\\r" },
                {"\t", "\\t" },
                {"\\u", "\\\\u" },
            };

        /// <summary>
        /// エスケープ文字をエスケープする
        /// </summary>
        /// <param name="text">エスケープ前の文字列</param>
        /// <returns>エスケープ後の文字列</returns>
        public string EscapeText(string text)
        {
            string ret = text;
            foreach (var dic in EscapeStringsDic)
            {
                ret = ret.Replace(dic.Key, dic.Value);
            }
            return ret;
        }
    }
}
