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
        private string GetTestDir()
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


        [TestMethod()]
        public void InputTest()
        {
            ZkExcelToJsonl test = new ZkExcelToJsonl();

            string dir = GetTestDir();

            string infile = Path.Combine(dir, BaseTestFile);
            test = new ZkExcelToJsonl(infile, true, 1, 1, 1, 0);         // 初期設定(ヘッダあり, 1行1列目スタート, チェック列1)
            test.Read();                                                 // 読み込み
            if (test.Rows.Count != 100) { Assert.Fail(); }               // 100行読み込む想定

            infile = Path.Combine(dir, "test_base_noheader_1_1.xlsx");
            test = new ZkExcelToJsonl(infile, false, 1, 1, 1, 0);        // 初期設定(ヘッダなし, 1行1列目スタート, チェック列1)
            test.Read();                                                 // 読み込み
            if (test.Rows.Count != 100) { Assert.Fail(); }               // 100行読み込む想定

            infile = Path.Combine(dir, "test_base_2_2.xlsx");
            test = new ZkExcelToJsonl(infile, true, 2, 2, 2, 0);          // 初期設定(ヘッダあり, 2行2列目スタート, チェック列2)
            test.Read();                                                  // 読み込み
            if (test.Rows.Count != 100) { Assert.Fail(); }                // 100行読み込む想定

            infile = Path.Combine(dir, "test_base_noheader_2_2.xlsx");
            test = new ZkExcelToJsonl(infile, false, 2, 2, 2, 0);        // 初期設定(ヘッダなし, 2行2列目スタート, チェック列2)
            test.Read();                                                 // 読み込み
            if (test.Rows.Count != 100) { Assert.Fail(); }               // 100行読み込む想定


            infile = Path.Combine(dir, "test_base_2_2_3.xlsx");
            test = new ZkExcelToJsonl(infile, true, 2, 2, 3, 0);         // 初期設定(ヘッダなし, 2行2列目スタート, チェック列3)
            test.Read();                                                 // 読み込み
            if (test.Rows.Count != 10) { Assert.Fail(); }                // 10行読み込む想定

            infile = Path.Combine(dir, "test_base_sheet2.xlsx");
            test = new ZkExcelToJsonl(infile, true, 1, 1, 1, 1);         // 初期設定(ヘッダなし, 2行2列目スタート, チェック列3, 2つ目のシートの読み込み)
            test.Read();                                                 // 読み込み
            if (test.Rows.Count != 50) { Assert.Fail(); }                // 50行読み込む想定
        }

        [TestMethod()]
        public void OutputTest()
        {

            string dir = GetTestDir();
            string infile = Path.Combine(dir, BaseTestFile);

            string outdir = Path.Combine(dir, "result");
            string filename = Path.GetFileNameWithoutExtension(infile) + ".jsonl";
            string outfile = Path.Combine(outdir, filename);

            PathManager.CreateDirectory(outdir);    // 再帰的にディレクトリの作成

            ZkExcelToJsonl test = new ZkExcelToJsonl(infile, true, 1, 1, 1, 0);
            test.Read();
            test.Write(outfile);
        }
    }
}