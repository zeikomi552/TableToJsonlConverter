using Microsoft.VisualStudio.TestTools.UnitTesting;
using TableToJsonlConverter.Conveters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;

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
        public void InputTest()
        {

            string dir = GetTestDir();
            string infile = Path.Combine(dir, BaseTestFile);

            ZkCsvToJsonl test = new(infile, Encoding.UTF8, true);
            test.Read();

            if (test.Rows.Count != 100) { Assert.Fail(); }
        }
    }
}