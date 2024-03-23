﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using DocumentFormat.OpenXml.Vml;
using Path = System.IO.Path;
using TableToJsonlConverter.Conveters;

namespace TableToJsonlConverterTests.Conveters
{
    [TestClass()]
    public class ZkExcelToJsonlTests
    {
        const string BaseTestFile = "test_base.xlsx";
        [TestMethod()]
        public void ExcelToJsonlTest()
        {
        }

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
        public void InitializeTest()
        {
            ZkExcelToJsonl test = new ZkExcelToJsonl();
            string dir = GetTestDir();

            string test_path = Path.Combine(dir, BaseTestFile);

            if (File.Exists(test_path))
            {
                string outfile = "test.json";

                test.Initialize(test_path, outfile, 1, 1, 1, 0);
                if (test.Initialize(test_path, outfile, 1, 1, 1, 0) != true) { Assert.Fail(); }

                if (test.Initialize("", outfile, 1, 1, 1, 0) == true) { Assert.Fail(); }

                if (test.Initialize(test_path, "", 1, 1, 1, 0) == true) { Assert.Fail(); }

                if (test.Initialize(test_path, outfile, 0, 1, 1, 0) == true) { Assert.Fail(); }

                if (test.Initialize(test_path, outfile, 1, 0, 1, 0) == true) { Assert.Fail(); }

                if (test.Initialize(test_path, outfile, 1, 1, 0, 0) == true) { Assert.Fail(); }

                if (test.Initialize(test_path, outfile, 1, 1, 1, -1) == true) { Assert.Fail(); }
            }
        }

        [TestMethod()]
        public void InputTest()
        {
            ZkExcelToJsonl test = new ZkExcelToJsonl();

            string dir = GetTestDir();
            string infile = Path.Combine(dir, BaseTestFile);
            string outfile = "test.json";

            if (test.Initialize(infile, outfile, 1, 1, 1, 0) == false) { Assert.Fail(); }

            test.Input();

            if (test.Rows.Count != 100) { Assert.Fail(); }

            infile = Path.Combine(dir, "test_base_noheader_1_1.xlsx");
            if (test.Initialize(infile, outfile, 1, 1, 1, 0, false) == false) { Assert.Fail(); }
            test.Input();
            if (test.Rows.Count != 100) { Assert.Fail(); }


            infile = Path.Combine(dir, "test_base_2_2.xlsx");
            if (test.Initialize(infile, outfile, 2, 2, 2, 0, true) == false) { Assert.Fail(); }
            test.Input();
            if (test.Rows.Count != 100) { Assert.Fail(); }

            infile = Path.Combine(dir, "test_base_noheader_2_2.xlsx");
            if (test.Initialize(infile, outfile, 2, 2, 2, 0, false) == false) { Assert.Fail(); }
            test.Input();
            if (test.Rows.Count != 100) { Assert.Fail(); }

            infile = Path.Combine(dir, "test_base_2_2_3.xlsx");
            if (test.Initialize(infile, outfile, 2, 2, 3, 0, true) == false) { Assert.Fail(); }
            test.Input();
            if (test.Rows.Count != 10) { Assert.Fail(); }

            infile = Path.Combine(dir, "test_base_sheet2.xlsx");
            if (test.Initialize(infile, outfile, 1, 1, 1, 1, true) == false) { Assert.Fail(); }
            test.Input();
            if (test.Rows.Count != 50) { Assert.Fail(); }
        }

        [TestMethod()]
        public void OutputTest()
        {
            ZkExcelToJsonl test = new ZkExcelToJsonl();

            string dir = GetTestDir();
            string infile = Path.Combine(dir, BaseTestFile);
            string outfile = "test.json";

            if (test.Initialize(infile, outfile, 1, 1, 1, 0) == false) { Assert.Fail(); }

            test.Input();

            test.Output();
        }
    }
}