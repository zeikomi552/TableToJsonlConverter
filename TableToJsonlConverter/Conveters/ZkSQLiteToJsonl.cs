﻿using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;
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
    /// SQLiteテーブルをJson Linesの変換を行うクラス
    /// </summary>
    public class ZkSQLiteToJsonl : ZkJsonlBase, IZkTableToJsonl
    {
        #region 入力ファイルパス(SQLite)
        public string ConnectionString { get; set; } = string.Empty;
        #endregion
        public string SQLCommand { get; set; } = string.Empty;

        public ZkSQLiteToJsonl()
        {

        }

        public ZkSQLiteToJsonl(string conectionString, string sql)
        {
            this.ConnectionString = conectionString;
            this.SQLCommand = sql;
        }

        #region ヘッダー情報の取得関数
        /// <summary>
        /// ヘッダー情報の取得関数
        /// </summary>
        /// <param name="row">指定した行のヘッダ情報(動的に変化する様なものでない限り1行目と変わらないので0を指定しておけばよい)</param>
        /// <returns>ヘッダー情報</returns>
        private ZkHeaders? GetHeader(SqliteDataReader sdr)
        {
            ZkHeaders header = new ZkHeaders();
            int field_max = sdr.FieldCount;

            for (int i = 0; i < field_max; i++)
            {
                string hd = sdr.GetName(i);
                header.Add(new ZkHeaderData() { Col = i, ColumnName = hd });
            }
            return header;
        }
        #endregion

        #region 入力処理
        /// <summary>
        /// 入力処理
        /// 失敗時はthrowを投げます
        /// </summary>
        public void Read(int read_rowcount = -1)
        {
            try
            {
                using (var connection = new SqliteConnection($"{this.ConnectionString}"))
                {
                    connection.Open();

                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = this.SQLCommand;

                        using (var reader = command.ExecuteReader())
                        {
                            ZkHeaders? header = null;
                            ZkRows rows = new ZkRows();

                            int row_cnt = 1;
                            while (reader.Read())
                            {
                                if (header == null)
                                {
                                    header = GetHeader(reader);

                                    if (header == null)
                                    {
                                        // ヘッダー情報未取得
                                        throw new Exception("Header information could not be retrieved");
                                    }
                                }

                                int field_max = reader.FieldCount;

                                ZkRow row = new ZkRow();
                                for (int i = 0; i < field_max; i++)
                                {
                                    ZkCellData cell = new ZkCellData() { Col = i, Key = header[i].ColumnName, Value = reader.GetValue(i) };
                                    row.Add(cell);
                                }
                                rows.Add(row);

                                row_cnt++;
                                if (read_rowcount > 0 && row_cnt > read_rowcount)
                                    break;

                            }
                            this.Rows = rows;
                        }
                    }
                    connection.Close();
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
