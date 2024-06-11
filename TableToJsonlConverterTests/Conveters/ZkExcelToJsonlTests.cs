using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using DocumentFormat.OpenXml.Vml;
using Path = System.IO.Path;
using TableToJsonlConverter.Conveters;
using TableToJsonlConverter.Utils;
using Microsoft.Data.SqlClient;
using TableToJsonlConverter.Conveters.Tests;

namespace TableToJsonlConverterTests.Conveters.Tests
{

    [TestClass()]
    public class ZkExcelToJsonlTests
    {
        [TestMethod()]
        public void ZkExcelToJsonlTest()
        {
            ZkExcelToJsonl test = new ZkExcelToJsonl();
            string dir = GetTestDir();

            string test_path = Path.Combine(dir, BaseTestFile);

            test = new ZkExcelToJsonl(test_path, true, 1, 1, 1, 0);
            if (!test.PropertyOk) { Assert.Fail(); }

            test = new ZkExcelToJsonl("", true, 1, 1, 1, 0);
            if (test.PropertyOk) { Assert.Fail(); }

            test = new ZkExcelToJsonl(test_path, true, 0, 1, 1, 0);
            if (test.PropertyOk) { Assert.Fail(); }

            test = new ZkExcelToJsonl(test_path, true, 1, 0, 1, 0);
            if (test.PropertyOk) { Assert.Fail(); }

            test = new ZkExcelToJsonl(test_path, true, 1, 1, 0, 0);
            if (test.PropertyOk) { Assert.Fail(); }

            test = new ZkExcelToJsonl(test_path, true, 1, 1, 1, -1);
            if (test.PropertyOk) { Assert.Fail(); }
        }

        const string BaseTestFile = "test_base.xlsx";

        #region テスト用のディレクトリファイルが置かれている場所を取得する
        /// <summary>
        /// テスト用のディレクトリファイルが置かれている場所を取得する
        /// </summary>
        /// <returns>テスト用ディレクトリパス</returns>
        public static string GetTestDir()
        {
            ZkExcelToJsonl test = new ZkExcelToJsonl();
            Assembly myAssembly = Assembly.GetEntryAssembly()!;
            string? path = myAssembly.Location;

            while (path != null)
            {
                path = Path.GetDirectoryName(path);

                if (string.IsNullOrEmpty(path))
                    break;

                string test_path = Path.Combine(path, "TestFiles", "Excel", BaseTestFile);

                if (File.Exists(test_path))
                {
                    return Path.Combine(path, "TestFiles", "Excel");
                }
            }

            return string.Empty;
        }
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="inpath">入力ファイルパス</param>
        /// <param name="headerf">入力ファイルにヘッダーの有無(true:ヘッダ有り false:ヘッダなし)</param>
        /// <param name="scol">データの開始列(1から始まる, ヘッダーを含む場合はヘッダーの開始位置)</param>
        /// <param name="srow">データの開始行(1から始まる, ヘッダーを含む場合はヘッダーの開始位置)</param>
        /// <param name="checkcol">必ずデータが入る列, ここの値が空になったらデータの取得をやめる</param>
        /// <param name="sheetno">入力ファイルのシート番号(0から始まる)</param>
        /// <returns></returns>
        private ZkExcelToJsonl ReadSub(string filename, bool headerf = true, int scol = 1, int srow = 1, int checkcol = 1, int sheetno = 0, int rowcount = -1)
        {
            ZkExcelToJsonl test = new ZkExcelToJsonl();

            string dir = GetTestDir();

            string infile = Path.Combine(dir, filename);

            test = new ZkExcelToJsonl()
            {
                InputPath = infile,             // input file path for .xlsx
                HeaderF = headerf,              // hedder -> true:exist false:not exist
                StartCol = scol,                // start column -> 1~
                StartRow = srow,                // start row -> 1~
                CheckCol = checkcol,            // not null column -> 1~
                SheetNo = sheetno               // excel sheet index -> 0~
            };

            test.Read(rowcount);

            return test;
        }

        [TestMethod()]
        public void ReadTest()
        {
            var test = ReadSub(BaseTestFile, true, 1, 1, 1, 0);;
            Assert.AreEqual(100, test.Rows.Count);

            test = ReadSub("test_base_noheader_1_1.xlsx", false, 1, 1, 1, 0);           // 初期設定(ヘッダなし, 1行1列目スタート, チェック列1)
            Assert.AreEqual(100, test.Rows.Count);

            test = ReadSub("test_base_2_2.xlsx", true, 2, 2, 2, 0);                     // 初期設定(ヘッダあり, 2行2列目スタート, チェック列2)
            Assert.AreEqual(100, test.Rows.Count);

            test = ReadSub("test_base_noheader_2_2.xlsx", false, 2, 2, 2, 0);           // 初期設定(ヘッダなし, 2行2列目スタート, チェック列2)
            Assert.AreEqual(100, test.Rows.Count);

            test = ReadSub("test_base_2_2_3.xlsx", true, 2, 2, 3, 0);                   // 初期設定(ヘッダなし, 2行2列目スタート, チェック列3)
            Assert.AreEqual(100, test.Rows.Count);

            test = ReadSub("test_base_sheet2.xlsx", true, 1, 1, 1, 1);                  // 初期設定(ヘッダなし, 2行2列目スタート, チェック列3, 2つ目のシートの読み込み)
            Assert.AreEqual(50, test.Rows.Count);


            test = ReadSub(BaseTestFile, true, 1, 1, 1, 0, 10); ;
            Assert.AreEqual(10, test.Rows.Count);

            test = ReadSub("test_base_noheader_1_1.xlsx", false, 1, 1, 1, 0, 10);           // 初期設定(ヘッダなし, 1行1列目スタート, チェック列1)
            Assert.AreEqual(10, test.Rows.Count);

            test = ReadSub("test_base_2_2.xlsx", true, 2, 2, 2, 0, 10);                     // 初期設定(ヘッダあり, 2行2列目スタート, チェック列2)
            Assert.AreEqual(10, test.Rows.Count);

            test = ReadSub("test_base_noheader_2_2.xlsx", false, 2, 2, 2, 0, 10);           // 初期設定(ヘッダなし, 2行2列目スタート, チェック列2)
            Assert.AreEqual(10, test.Rows.Count);

            test = ReadSub("test_base_2_2_3.xlsx", true, 2, 2, 3, 0, 10);                   // 初期設定(ヘッダなし, 2行2列目スタート, チェック列3)
            Assert.AreEqual(10, test.Rows.Count);

            test = ReadSub("test_base_sheet2.xlsx", true, 1, 1, 1, 1, 10);                  // 初期設定(ヘッダなし, 2行2列目スタート, チェック列3, 2つ目のシートの読み込み)
            Assert.AreEqual(10, test.Rows.Count);
        }

        [TestMethod()]
        public void WriteTest()
        {
            var test = ReadSub(BaseTestFile, true, 1, 1, 1, 0);

            string file_path = "excel.json";
            var base_test_dir = ZkJsonlBaseTests.GetTestBaseDir();
            var test_dir = Path.Combine(base_test_dir, "Excel", "results");
            var filepath = Path.Combine(test_dir, file_path);
            PathManager.CreateCurrentDirectory(filepath);

            var filepath_gz = Path.Combine(test_dir, file_path);
            test.Write(filepath);
            test.CompressWrite(filepath_gz);
        }

    }
}