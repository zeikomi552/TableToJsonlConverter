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

                string test_path = Path.Combine(path, "testfile", BaseTestFile);

                if (File.Exists(test_path))
                {
                    return Path.Combine(path, "testfile");
                }
            }

            return string.Empty;
        }
        [TestMethod()]
        public void InputTest()
        {
            ZkCsvToJsonl test = new ZkCsvToJsonl();

            string dir = GetTestDir();
            string infile = Path.Combine(dir, BaseTestFile);
            string outfile = "test_csv.json";

            if (test.Initialize(infile, outfile, Encoding.UTF8, true) == false) { Assert.Fail(); }

            test.Read();

            if (test.Rows.Count != 100) { Assert.Fail(); }

            test.Write(outfile);
        }
    }
}