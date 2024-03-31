using Microsoft.VisualStudio.TestTools.UnitTesting;
using TableToJsonlConverter.Conveters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using TableToJsonlConverter.Utils;
using DocumentFormat.OpenXml.Math;

namespace TableToJsonlConverter.Conveters.Tests
{
    [TestClass()]
    public class ZkCsvToJsonlTests
    {
        const string BaseTestFile = "test_base.csv";

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

                string test_path = Path.Combine(path, "TestFiles", "Csv", BaseTestFile);

                if (File.Exists(test_path))
                {
                    return Path.Combine(path, "TestFiles", "Csv");
                }
            }

            return string.Empty;
        }
        #endregion
        [TestMethod()]
        public void ReadTest()
        {
            var test = ReadSub();

            if (test.Rows.Count != 100) { Assert.Fail(); }
        }

        private ZkCsvToJsonl ReadSub()
        {
            string dir = GetTestDir();
            string infile = Path.Combine(dir, BaseTestFile);

            ZkCsvToJsonl test = new(infile, Encoding.UTF8, true);
            test.Read();

            return test;
        }
        private ZkCsvToJsonl ReadSub2()
        {
            string dir = GetTestDir();
            string infile = Path.Combine(dir, "test_base.tsv");

            ZkCsvToJsonl test = new(infile, Encoding.UTF8, true,"\t");
            test.Read();

            return test;
        }
        [TestMethod()]
        public void WriteTest()
        {
            var read_result = ReadSub();

            string file_path = "csv_test.json";
            var base_test_dir = ZkJsonlBaseTests.GetTestBaseDir();
            var test_dir = Path.Combine(base_test_dir, "csv", "results");
            var filepath = Path.Combine(test_dir, file_path);
            PathManager.CreateCurrentDirectory(filepath);

            var filepath_gz = Path.Combine(test_dir, file_path);
            read_result.Write(filepath);
            read_result.CompressWrite(filepath_gz);


            read_result = ReadSub2();

            file_path = "csv_test.tsv.json";
            base_test_dir = ZkJsonlBaseTests.GetTestBaseDir();
            test_dir = Path.Combine(base_test_dir, "csv", "results");
            filepath = Path.Combine(test_dir, file_path);
            PathManager.CreateCurrentDirectory(filepath);

            filepath_gz = Path.Combine(test_dir, file_path);
            read_result.Write(filepath);
            read_result.CompressWrite(filepath_gz);

        }
    }
}