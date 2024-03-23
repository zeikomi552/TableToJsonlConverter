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
        public virtual string EscapeText(string text)
        {
            string ret = text;
            foreach (var dic in EscapeStringsDic)
            {
                ret = ret.Replace(dic.Key, dic.Value);
            }
            return ret;
        }
        #endregion

        #region 出力ファイルパス
        /// <summary>
        /// 出力ファイルパス
        /// </summary>
        public string OutputPath { get; protected set; } = string.Empty;
        #endregion

        #region JsonLines
        /// <summary>
        /// JsonLines
        /// </summary>
        public virtual string JsonLines
        {
            get
            {
                StringBuilder jsonl = new StringBuilder();
                for (int i = 0; i < Rows.Count; i++)
                {
                    var row = Rows.ElementAt(i);
                    jsonl.Append("{");

                    for (int j = 0; j < row.Count; j++)
                    {
                        var item = row[j];
                        jsonl.Append($"\"{EscapeText(item.Key)}\": \"{EscapeText(item.Value.ToString()!)}\"");

                        if (j != row.Count - 1)
                            jsonl.Append(",");
                    }
                    jsonl.Append("}");

                    if (i != Rows.Count - 1)
                    {
                        jsonl.Append("\n");
                    }
                }
                return jsonl.ToString();
            }
        }
        #endregion

        #region 出力処理
        /// <summary>
        /// JsonLines 出力処理
        /// 失敗時はthrowを投げます
        /// </summary>
        /// <param name="filepath">ファイルパス</param>
        public virtual void Output()
        {
            try
            {
                File.WriteAllText(OutputPath, JsonLines);
            }
            catch
            {
                throw;
            }
        }
        #endregion
    }
}
