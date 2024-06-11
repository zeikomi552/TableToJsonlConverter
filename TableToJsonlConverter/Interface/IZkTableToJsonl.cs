using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableToJsonlConverter.Interface
{
    /// <summary>
    /// 各種JsonLines変換クラスのインターフェースクラス
    /// </summary>
    public interface IZkTableToJsonl
    {
        /// <summary>
        /// 対象データの読み込み処理
        /// </summary>
        void Read(int read_rowcount = -1);

        /// <summary>
        /// Json Linesファイルの書き出し処理
        /// </summary>
        /// <param name="path"></param>
        void Write(string path);

        string JsonLines { get; }
        /// <summary>
        /// GZip形式の圧縮付き書き出し処理(.gz)
        /// </summary>
        void CompressWrite();
    }
}
