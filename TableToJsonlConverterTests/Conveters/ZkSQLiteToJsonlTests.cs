using Microsoft.VisualStudio.TestTools.UnitTesting;
using TableToJsonlConverter.Conveters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;
using TableToJsonlConverter.Models;
using TableToJsonlConverter.Utils;

namespace TableToJsonlConverter.Conveters.Tests
{

    [TestClass()]
    public class ZkSQLiteToJsonlTests
    {
#warning Modifications are required for the test environment.
        const string testdb = "test.db";
        const string Connectionstring = $"Data Source={testdb}";
        const string Sqlcommand = "select * From [TEST]";

        private void CreateSQLite()
        {
            if (File.Exists(testdb))
            {
                File.Delete(testdb);
            }

            var connection = new SqliteConnection($"{Connectionstring}");
            {
                connection.Open();

                //テーブル作成
                using (var cmd = connection.CreateCommand())
                {
                    cmd.CommandText = "CREATE TABLE IF NOT EXISTS [TEST](" +
                                      "[id] INTEGER PRIMARY KEY," +
                                      "[name] TEXT NOT NULL" +
                                      ")";
                    cmd.ExecuteNonQuery();
                }
                for (int i = 0; i < 100; i++)
                {
                    //INSERT(1)
                    using (var cmd = connection.CreateCommand())
                    {

                        cmd.CommandText = "INSERT INTO [TEST]([name]) VALUES (@name)";
                        cmd.Parameters.Add(new Microsoft.Data.Sqlite.SqliteParameter() { ParameterName = "@name", Value = $"TEST{i}" });
                        cmd.ExecuteNonQuery();
                    }
                }
                connection.Close();
            }
        }

        [TestMethod()]
        public void ReadTest()
        {
            var test = ReadSub();

            Assert.AreEqual(100, test.Rows.Count);
            Assert.AreEqual(2, test.GetHeader(0)!.Count);

            test = ReadSub(false, 10);

            Assert.AreEqual(10, test.Rows.Count);
            Assert.AreEqual(2, test.GetHeader(0)!.Count);
        }

        private ZkSQLiteToJsonl ReadSub(bool create_f = true, int readcount = -1)
        {
            if (create_f)
            {
                CreateSQLite();
            }

            ZkSQLiteToJsonl tojson = new ZkSQLiteToJsonl()
            {
                ConnectionString = Connectionstring,
                SQLCommand = Sqlcommand
            };
            tojson.Read(readcount);

            return tojson;
        }


        [TestMethod()]
        public void WriteTest()
        {
            var test = ReadSub(false);

            string file_path = "sqlite.json";
            var base_test_dir = ZkJsonlBaseTests.GetTestBaseDir();
            var test_dir = Path.Combine(base_test_dir, "Sqlite", "results");
            var filepath = Path.Combine(test_dir, file_path);
            PathManager.CreateCurrentDirectory(filepath);

            var filepath_gz = Path.Combine(test_dir, file_path);
            test.Write(filepath);
            test.CompressWrite(filepath_gz);
        }
    }
}