﻿using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableToJsonlConverter.Interface;
using TableToJsonlConverter.Models;

namespace TableToJsonlConverter.Conveters
{
    /// <summary>
    /// CSVファイルをJson Linesの変換を行うクラス
    /// </summary>
    public class ZkCsvToJsonl : ZkJsonlBase, IZkTableToJsonl
    {
        #region 入力ファイルパス(Csv)
        /// <summary>
        /// 入力ファイルパス(Csv)
        /// </summary>
        public string InputPath { get; set; } = string.Empty;
        #endregion

        #region ヘッダーが存在するかどうか(true:ヘッダあり false:ヘッダなし) 
        /// <summary>
        /// ヘッダーが存在するかどうか(true:ヘッダあり false:ヘッダなし) 
        /// </summary>
        public bool HeaderF { get; set; } = true;

        /// <summary>
        /// ファイルのエンコーディング
        /// </summary>
        public Encoding Encoding { get; set; } = Encoding.UTF8;
        /// <summary>
        /// 区切り文字
        /// </summary>
        public string Delimiter { get; set; } = ",";
        #endregion

        #region コンストラクタ
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ZkCsvToJsonl()
        {

        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="inpath">入力ファイルパス</param>
        /// <param name="enc">エンコーディング</param>
        /// <param name="headerf">入力ファイルにヘッダーの有無(true:ヘッダ有り false:ヘッダなし)</param>
        public ZkCsvToJsonl(string ipath, Encoding enc, bool headerf = true, string delimiter = ",")
        {
            Rows = new ZkRows();
            InputPath = ipath;
            this.HeaderF = headerf;

            this.Encoding = enc;
            this.Delimiter = delimiter;
        }
        #endregion

        #region パラメータのチェックプロパティ
        /// <summary>
        /// パラメータのチェックプロパティ
        /// </summary>
        /// <returns>true:パラメータに不整値あり false:OK</returns>
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

                return isok;
            }
        }
        #endregion

        #region 読み込み処理
        /// <summary>
        /// 対象データの読み込み処理
        /// </summary>
        public void Read(int read_rowcount = -1)
        {
            try
            {
                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    //読み取ったヘッダが小文字に変換されるように ToLower() を仕込みます。
                    PrepareHeaderForMatch = args => args.Header.ToLower(),
                    Delimiter= this.Delimiter
                };

                int row_cnt = 1;

                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using (var reader = new StreamReader(this.InputPath, this.Encoding))
                using (var csv = new CsvHelper.CsvReader(reader, config))
                {
                    while (csv.Read())
                    {
                        var columns = csv.GetRecord<dynamic>() as IDictionary<string, object>;

                        if (columns == null)
                            return;

                        var row_tmp = new List<KeyValuePair<string, object>>();

                        int col = 1;
                        ZkRow row = new ZkRow();
                        foreach (var column in columns)
                        {
                            var data = new ZkCellData() { Col = col, Key = column.Key, Value = column.Value};
                            row.Add(data);
                            col++;
                        }

                        this.Rows.Add(row);

                        row_cnt++;
                        if (read_rowcount > 0 && row_cnt > read_rowcount)
                            break;
                    }
                }
            }
            catch
            {
                throw;
            }
        }
        #endregion

    }
}
