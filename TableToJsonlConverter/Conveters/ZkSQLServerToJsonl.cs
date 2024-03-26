using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Data.SqlClient;
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
    /// SQL Server Table to JsonLines
    /// </summary>
    public class ZkSQLServerToJsonl : ZkJsonlBase, IZkTableToJsonl
    {
        public string ConnectionString { get; set; } = string.Empty;
        public string SQLCommand { get; set; } = string.Empty;

        #region コンストラクタ
        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ZkSQLServerToJsonl()
        {

        }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="connection_string">コネクションストリング文字列</param>
        /// <param name="sqlcommand">実行するSQL</param>
        public ZkSQLServerToJsonl(string connection_string, string sqlcommand)
        {
            this.ConnectionString = connection_string;
            this.SQLCommand = sqlcommand;
        }
        #endregion

        #region ヘッダー情報の取得関数
        /// <summary>
        /// ヘッダー情報の取得関数
        /// </summary>
        /// <param name="row">指定した行のヘッダ情報(動的に変化する様なものでない限り1行目と変わらないので0を指定しておけばよい)</param>
        /// <returns>ヘッダー情報</returns>
        private ZkHeaders? GetHeader(SqlDataReader sdr)
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
        public void Read()
        {
            try
            {
                string constr = this.ConnectionString;

                SqlConnection con = new SqlConnection(constr);
                con.Open();

                try
                {
                    SqlCommand com = new SqlCommand(this.SQLCommand, con);
                    SqlDataReader sdr = com.ExecuteReader();

                    ZkHeaders? header = null;
                    ZkRows rows = new ZkRows();
                    while (sdr.Read() == true)
                    {
                        int field_max= sdr.FieldCount;

                        if (header == null)
                        {
                            header = GetHeader(sdr);

                            if (header == null)
                            {
                                // ヘッダー情報未取得
                                throw new Exception("Header information could not be retrieved");
                            }
                        }


                        ZkRow row = new ZkRow();
                        for (int i = 0; i < field_max; i++)
                        {
                            ZkCellData cell = new ZkCellData() { Col = i, Key = header[i].ColumnName, Value = sdr[i] };
                            row.Add(cell);
                        }
                        rows.Add(row);
                    }
                    this.Rows = rows;
                    sdr.Close();
                    com.Dispose();
                }
                finally
                {
                    con.Close();
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
