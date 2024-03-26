using Microsoft.VisualStudio.TestTools.UnitTesting;
using TableToJsonlConverter.Conveters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;
using TableToJsonlConverter.Models;

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
            CreateSQLite();
            ZkSQLiteToJsonl tojson = new ZkSQLiteToJsonl()
            {
                ConnectionString = Connectionstring,
                SQLCommand = Sqlcommand
            };
            tojson.Read();

            if (tojson.Rows.Count != 100)
            {
                Assert.Fail();
            }

            if (tojson.GetHeader(0)!.Count != 2)
            {
                Assert.Fail();
            }
        }
    }
}