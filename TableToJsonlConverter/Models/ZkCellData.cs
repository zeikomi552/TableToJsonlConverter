using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableToJsonlConverter.Models
{
    public class ZkCellData
    {
        /// <summary>
        /// カラム(列)番号 1から始まる
        /// </summary>
        public int Col { get; set; } = -1;

        /// <summary>
        /// キー
        /// </summary>
        public string Key { get; set; } = string.Empty;

        /// <summary>
        /// 値
        /// </summary>
        public object? Value { get; set; } = null;

        /// <summary>
        /// 型情報
        /// </summary>
        public Type? Type
        {
            get
            {
                if (Value != null)
                {
                    return Value.GetType();
                }
                else { return null; }
            }
        }
    }
}
