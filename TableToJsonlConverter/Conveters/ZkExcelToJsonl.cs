using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;
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
        public string InputPath { get; set; } = string.Empty;
        #endregion

        #region 開始カラム(1から始まる)
        /// <summary>
        /// 開始カラム(1から始まる)
        /// </summary>
        public int StartCol { get; set; } = 1;
        #endregion

        #region 開始行(1から始まる)
        /// <summary>
        /// 開始行(1から始まる)
        /// </summary>
        public int StartRow { get; set; } = 1;
        #endregion

        #region チェックするカラム（必ず値が入ることが前提で、入っていないセルを見つけるとそこで行を終了する）
        /// <summary>
        /// チェックするカラム（必ず値が入ることが前提で、入っていないセルを見つけるとそこで行を終了する）
        /// </summary>
        public int CheckCol { get; set; } = 1;
        #endregion

        #region 読み込むExcelのシート番号(0から始まる)
        /// <summary>
        /// 読み込むExcelのシート番号(0から始まる)
        /// </summary>
        public int SheetNo { get; set; } = 0;
        #endregion

        #region ヘッダーが存在するかどうか(true:ヘッダあり false:ヘッダなし) 
        /// <summary>
        /// ヘッダーが存在するかどうか(true:ヘッダあり false:ヘッダなし) 
        /// </summary>
        public bool HeaderF { get; set; } = true;
        #endregion
        #endregion

        #region コンストラクタ
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ZkExcelToJsonl()
        {

        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="inpath">入力ファイルパス</param>
        /// <param name="headerf">入力ファイルにヘッダーの有無(true:ヘッダ有り false:ヘッダなし)</param>
        /// <param name="scol">データの開始列(1から始まる, ヘッダーを含む場合はヘッダーの開始位置)</param>
        /// <param name="srow">データの開始行(1から始まる, ヘッダーを含む場合はヘッダーの開始位置)</param>
        /// <param name="checkcol">必ずデータが入る列, ここの値が空になったらデータの取得をやめる</param>
        /// <param name="sheetno">入力ファイルのシート番号(0から始まる)</param>
        public ZkExcelToJsonl(string inpath, bool headerf = true, int scol = 1, int srow = 1, int checkcol = 1, int sheetno=0)
        {
            InputPath = inpath;
            StartCol = scol;
            StartRow = srow;
            CheckCol = checkcol;
            SheetNo = sheetno;
            HeaderF = headerf;
        }
        #endregion

        #region パラメータのチェック処理
        /// <summary>
        /// プロパティに問題がないかのチェックを行う
        /// </summary>
        /// <returns>true:プロパティ(入力ファイルパスが存在しない, 開始位置が0以下, シート番号が0未満) false:OK</returns>
        public bool PropertyOk
        {
            get
            {
                bool isok = true;

                // 入力ファイルパス
                if (!File.Exists(InputPath))
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
        public void Read()
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
