using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableToJsonlConverter.Models;

namespace TableToJsonlConverter.Conveters
{
    public class ZkJsonlBase
    {
        #region ヘッダー情報
        /// <summary>
        /// ヘッダー情報
        /// </summary>
        public ZkHeaders Headers { get; protected set; } = new ZkHeaders();
        #endregion

        #region 行情報
        /// <summary>
        /// 行情報
        /// </summary>
        public ZkRows Rows { get; protected set; } = new ZkRows();
        #endregion

        #region エスケープ文字辞書
        /// <summary>
        /// エスケープ文字辞書
        /// </summary>
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
        #endregion

        #region エスケープ文字をエスケープする
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
        #endregion
    }
}
