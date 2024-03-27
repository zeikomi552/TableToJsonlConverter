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
        const string Connectionstring = @"Data Source=TANE2\SQLEXPRESS;Initial Catalog=AdventureWorks2019;User ID=testuser;Password=testuser;Encrypt=false;TrustServerCertificate=true";
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
            ZkSQLServerToJsonl zkSQLServerToJsonl = new ZkSQLServerToJsonl()
            {
                ConnectionString = Connectionstring,
                SQLCommand = Sqlcommand
            };

            zkSQLServerToJsonl.Read();

            if (zkSQLServerToJsonl.Rows.Count == 0)
            {
                Assert.Fail();
            }

            var base_test_dir = ZkJsonlBaseTests.GetTestBaseDir();
            var test_dir  = Path.Combine(base_test_dir, "SQLServer", "results");
            var filepath = Path.Combine(test_dir, "sqlserver.json");
            var filepath_gz = Path.Combine(test_dir, "sqlserver.json");

            PathManager.CreateCurrentDirectory(filepath);
            zkSQLServerToJsonl.Write(filepath);
            zkSQLServerToJsonl.CompressWrite(filepath_gz);


        }
    }
}