using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableToJsonlConverter.Interface;
using TableToJsonlConverter.Models;

namespace TableToJsonlConverter.Conveters
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
            Headers = new ZkHeaders();
            Rows = new ZkRows();
            InputPath = ipath;
            OutputPath = opath;
            StartCol = scol;
            StartRow = srow;
            CheckCol = chcol;
            SheetNo = sheetno;
            HeaderF = headerf;

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
            if (!File.Exists(InputPath))
            {
                isok = false;
            }

            // 出力ファイルパス
            if (string.IsNullOrEmpty(OutputPath))
            {
                isok = false;
            }

            // 各種行、列は1異以上が正常で、シート番号は0以上が正常
            if (StartCol <= 0 || StartRow <= 0 || CheckCol <= 0 || SheetNo < 0)
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
        /// 失敗時はthrowを投げます
        /// </summary>
        public void Input()
        {
            try
            {
                // 読み込み専用で開く
                using (FileStream fs = new FileStream(InputPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (XLWorkbook workbook = new XLWorkbook(fs))
                    {
                        var ws = workbook.Worksheets.ElementAt(SheetNo);

                        // ヘッダーのセット処理
                        SetHeader(ws);

                        // ヘッダーのセット処理
                        SetRows(ws);
                    }
                }
            }
            catch
            {
                throw;
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
            int col = StartCol;

            if (HeaderF)
            {
                while (true)
                {
                    var val = ws.Cell(StartRow, col).CachedValue;

                    if (string.IsNullOrEmpty(val.ToString()))
                    {
                        break;
                    }
                    else
                    {
                        Headers.Add(col, val.ToString());
                    }
                    col++;
                }
            }
            else
            {
                while (true)
                {
                    var val = ws.Cell(StartRow, col).CachedValue;

                    if (string.IsNullOrEmpty(val.ToString()))
                    {
                        break;
                    }
                    else
                    {
                        Headers.Add(col, "col" + col.ToString());
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
            int col = StartCol;
            int row = StartRow;

            // ヘッダが存在するので一行ずらす
            if (HeaderF)
                row++;

            while (true)
            {
                var row_tmp = new List<KeyValuePair<string, object>>();
                var check_cell = ws.Cell(row, CheckCol).CachedValue.ToString();

                // 対象セルが空白ならば抜ける
                if (string.IsNullOrEmpty(check_cell))
                {
                    break;
                }

                // ヘッダの数だけ回す
                foreach (var header in Headers)
                {
                    var val = ws.Cell(row, header.Key).CachedValue;
                    row_tmp.Add(new KeyValuePair<string, object>(header.Value, val));
                }

                Rows.Add(row_tmp);
                row++;
            }
        }
        #endregion
    }
}
