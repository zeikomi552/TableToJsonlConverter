using Microsoft.VisualStudio.TestTools.UnitTesting;
using TableToJsonlConverter.Conveters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using TableToJsonlConverterTests.Conveters.Tests;
using TableToJsonlConverter.Utils;

namespace TableToJsonlConverter.Conveters.Tests
{
    [TestClass()]
    public class ZkSQLServerToJsonlTests
    {
#warning Modifications are required for the test environment.
        const string Connectionstring = @"Data Source=TANE2\SQLEXPRESS;Initial Catalog=AdventureWorks2019;User ID=testuser;Password=testuser;Encrypt=False;Trust Server Certificate=True";
        const string Sqlcommand = "select top 100 * From Person.Person";


        [TestMethod()]
        public void ZkSQLServerToJsonlTest()
        {
            ZkSQLServerToJsonl zkSQLServerToJsonl = new ZkSQLServerToJsonl()
            {
                ConnectionString = Connectionstring,
                SQLCommand = Sqlcommand
            };

            //Assert.Fail();
        }

        [TestMethod()]
        public void ZkSQLServerToJsonlTest1()
        {
            ZkSQLServerToJsonl zkSQLServerToJsonl = new ZkSQLServerToJsonl()
            {
                ConnectionString = Connectionstring,
                SQLCommand = Sqlcommand
            };
            //Assert.Fail();
        }

        [TestMethod()]
        public void ReadTest()
        {
            var zkSQLServerToJsonl = ReadSub();

            if (zkSQLServerToJsonl.Rows.Count == 0)
            {
                Assert.Fail();
            }
        }

        private ZkSQLServerToJsonl ReadSub()
        {
            ZkSQLServerToJsonl zkSQLServerToJsonl = new ZkSQLServerToJsonl()
            {
                ConnectionString = Connectionstring,
                SQLCommand = Sqlcommand
            };

            zkSQLServerToJsonl.Read();

            return zkSQLServerToJsonl;
        }

        [TestMethod()]
        public void WriteTest()
        {
            var zkSQLServerToJsonl = ReadSub();

            string file_path = "sqlserver.json";
            var base_test_dir = ZkJsonlBaseTests.GetTestBaseDir();
            var test_dir = Path.Combine(base_test_dir, "SQLServer", "results");
            var filepath = Path.Combine(test_dir, file_path);
            PathManager.CreateCurrentDirectory(filepath);

            var filepath_gz = Path.Combine(test_dir, file_path);
            zkSQLServerToJsonl.Write(filepath);
            zkSQLServerToJsonl.CompressWrite(filepath_gz);

        }
    }
}