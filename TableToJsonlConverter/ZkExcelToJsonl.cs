using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableToJsonlConverter
{
    /// <summary>
    /// 表形式で表現されるエクセルをJson Lines形式に変換するクラス
    /// </summary>
    public class ZkExcelToJsonl : ZkJsonlBase, IZkTableToJsonl
    {
        #region Properties
        #region 入力ファイルパス(Excel)
        /// <summary>
        /// 入力ファイルパス(Excel)
        /// </summary>
        public string InputPath { get; private set; } = string.Empty;
        #endregion

        #region 出力ファイルパス
        /// <summary>
        /// 出力ファイルパス
        /// </summary>
        public string OutputPath { get; private set; } = string.Empty;
        #endregion

        #region 開始カラム(1から始まる)
        /// <summary>
        /// 開始カラム(1から始まる)
        /// </summary>
        public int StartCol { get; private set; } = 1;
        #endregion

        #region 開始行(1から始まる)
        /// <summary>
        /// 開始行(1から始まる)
        /// </summary>
        public int StartRow { get; private set; } = 1;
        #endregion

        #region チェックするカラム（必ず値が入ることが前提で、入っていないセルを見つけるとそこで行を終了する）
        /// <summary>
        /// チェックするカラム（必ず値が入ることが前提で、入っていないセルを見つけるとそこで行を終了する）
        /// </summary>
        public int CheckCol { get; private set; } = 1;
        #endregion

        #region 読み込むExcelのシート番号(0から始まる)
        /// <summary>
        /// 読み込むExcelのシート番号(0から始まる)
        /// </summary>
        public int SheetNo { get; private set; } = 0;
        #endregion

        #region ヘッダーが存在するかどうか(true:ヘッダあり false:ヘッダなし) 
        /// <summary>
        /// ヘッダーが存在するかどうか(true:ヘッダあり false:ヘッダなし) 
        /// </summary>
        public bool HeaderF { get; private set; } = true;
        #endregion

        #region ヘッダー情報
        /// <summary>
        /// ヘッダー情報
        /// </summary>
        public ZkHeaders Headers { get; private set; } = new ZkHeaders();
        #endregion

        #region 行情報
        /// <summary>
        /// 行情報
        /// </summary>
        public ZkRows Rows { get; private set; } = new ZkRows();
        #endregion

        #region JsonLines
        /// <summary>
        /// JsonLines
        /// </summary>
        public string JsonLines
        {
            get
            {
                StringBuilder jsonl = new StringBuilder();
                for (int i = 0; i < this.Rows.Count; i++)
                {
                    var row = this.Rows.ElementAt(i);
                    jsonl.Append("{");

                    for (int j = 0; j < row.Count; j++)
                    {
                        var item = row[j];
                        jsonl.Append($"\"{EscapeText(item.Key)}\": \"{EscapeText(item.Value.ToString()!)}\"");

                        if (j != row.Count - 1)
                            jsonl.Append(",");
                    }
                    jsonl.Append("}");

                    if (i != this.Rows.Count - 1)
                    {
                        jsonl.Append("\n");
                    }
                }
                return jsonl.ToString();
            }
        }
        #endregion
        #endregion

        #region コンストラクタ
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ZkExcelToJsonl()
        {

        }
        #endregion

        #region 初期化処理
        /// <summary>
        /// 初期化処理
        /// </summary>
        /// <param name="ipath">入力ファイルパス(.xlsx)</param>
        /// <param name="opath">出力ファイルパス(.jsonl)</param>
        /// <param name="scol">カラム開始位置(A列 = 1, 1から始まる)</param>
        /// <param name="srow">行開始位置(1から始まる)</param>
        /// <param name="chcol">チェックするカラム(必ず値が入っているカラムとし、空のセルを見つけた時点でデータ終了とする)</param>
        /// <param name="sheetno">読み込むExcelのシート番号(0から始まる)</param>
        /// <param name="headerf">true:ヘッダーあり false:ヘッダーなし</param>
        /// <returns>true:各設定値が正常 false:設定値が異常</returns>
        public bool Initialize(string ipath, string opath, int scol = 1, int srow = 1, int chcol = 1, int sheetno = 0, bool headerf = true)
        {
            this.Headers = new ZkHeaders();
            this.Rows = new ZkRows();
            this.InputPath = ipath;
            this.OutputPath = opath;
            this.StartCol = scol;
            this.StartRow = srow;
            this.CheckCol = chcol;
            this.SheetNo = sheetno;
            this.HeaderF = headerf;

            // パラメータのチェック
            return CheckParameter();
        }
        #endregion

        #region パラメータのチェック処理
        /// <summary>
        /// パラメータのチェック処理
        /// </summary>
        /// <returns>true:パラメータに不整値あり false:OK</returns>
        private bool CheckParameter()
        {
            bool isok = true;

            // 入力ファイルパス
            if (!File.Exists(this.InputPath))
            {
                isok = false;
            }

            // 出力ファイルパス
            if (string.IsNullOrEmpty(this.OutputPath))
            {
                isok = false;
            }

            // 各種行、列は1異以上が正常で、シート番号は0以上が正常
            if (this.StartCol <= 0 || this.StartRow <= 0 || this.CheckCol <= 0 || this.SheetNo < 0)
            {
                isok = false;
            }
            return isok;
        }
        #endregion

        #region パスの存在確認
        /// <summary>
        /// パスの存在確認
        /// </summary>
        /// <param name="path">ファイルパス</param>
        /// <returns>存在しない場合はstring.Empty 存在する場合は値をそのまま返却</returns>
        private string Checkpath(string path)
        {
            if (File.Exists(path))
            {
                return path;
            }
            else
            {
                return string.Empty;
            }
        }
        #endregion

        #region 入力処理
        /// <summary>
        /// 入力処理
        /// </summary>
        public void Input()
        {
            // 読み込み専用で開く
            using (FileStream fs = new FileStream(this.InputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (XLWorkbook workbook = new XLWorkbook(fs))
                {
                    var ws = workbook.Worksheets.ElementAt(this.SheetNo);

                    // ヘッダーのセット処理
                    SetHeader(ws);

                    // ヘッダーのセット処理
                    SetRows(ws);
                }
            }
        }
        #endregion

        #region ヘッダーのセット処理
        /// <summary>
        /// ヘッダーのセット処理
        /// </summary>
        /// <param name="ws">ワークシート</param>
        private void SetHeader(IXLWorksheet ws)
        {
            int col = this.StartCol;

            if (this.HeaderF)
            {
                while (true)
                {
                    var val = ws.Cell(this.StartRow, col).CachedValue;

                    if (string.IsNullOrEmpty(val.ToString()))
                    {
                        break;
                    }
                    else
                    {
                        this.Headers.Add(col, val.ToString());
                    }
                    col++;
                }
            }
            else
            {
                while (true)
                {
                    var val = ws.Cell(this.StartRow, col).CachedValue;

                    if (string.IsNullOrEmpty(val.ToString()))
                    {
                        break;
                    }
                    else
                    {
                        this.Headers.Add(col, "col" + col.ToString());
                    }
                    col++;
                }
            }
        }
        #endregion

        #region 行のセット処理
        /// <summary>
        /// 行のセット処理
        /// </summary>
        /// <param name="ws">エクセルワークシート</param>
        private void SetRows(IXLWorksheet ws)
        {
            int col = this.StartCol;
            int row = this.StartRow;

            // ヘッダが存在するので一行ずらす
            if (this.HeaderF)
                row++;

            while (true)
            {
                var row_tmp = new List<KeyValuePair<string, object>>();
                var check_cell = ws.Cell(row, this.CheckCol).CachedValue.ToString();

                // 対象セルが空白ならば抜ける
                if (string.IsNullOrEmpty(check_cell))
                {
                    break;
                }

                // ヘッダの数だけ回す
                foreach (var header in this.Headers)
                {
                    var val = ws.Cell(row, header.Key).CachedValue;
                    row_tmp.Add(new KeyValuePair<string, object>(header.Value, val));
                }

                this.Rows.Add(row_tmp);
                row++;
            }
        }
        #endregion

        #region 出力処理
        /// <summary>
        /// JsonLines 出力処理
        /// </summary>
        /// <param name="filepath">ファイルパス</param>
        public void Output()
        {
            File.WriteAllText(this.OutputPath, this.JsonLines);
        }
        #endregion
    }
}
