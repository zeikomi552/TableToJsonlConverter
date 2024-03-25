using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using TableToJsonlConverter.Conveters;
using TableToJsonlConverterTests.Conveters.Tests;

namespace TableToJsonlConverter.Conveters.Tests
{
    [TestClass()]
    public class ZkJsonlBaseTests
    {
        const string BaseTestFile = "test_base.xlsx";

        public static string GetTestBaseDir()
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
                    return Path.Combine(path, "TestFiles");
                }
            }

            return string.Empty;
        }

        [TestMethod()]
        public void GetHeaderTest()
        {
            ZkExcelToJsonl test = new ZkExcelToJsonl();
            string dir = ZkExcelToJsonlTests.GetTestDir();

            string test_path = Path.Combine(dir, BaseTestFile);

            test = new ZkExcelToJsonl(test_path, true, 1, 1, 1, 0);
            if (!test.PropertyOk) { Assert.Fail(); }
            test.Read();
            var header = test.GetHeader(0);

            if (header!.Count != 4) { Assert.Fail(); }

            int i = 1;
            foreach (var tmp in header)
            {
                string column_name = $"header{i++}";
                var find = (from x in header
                            where x.ColumnName.Contains(column_name)
                            select x).Any();

                if (!find)
                {
                    Assert.Fail();
                }
            }
        }

        [TestMethod()]
        public void EscapeTextTest()
        {
            ZkJsonlBase test = new ZkJsonlBase();

            for (char ch = (char)0; ch < (char)255; ch++)
            {
                string ans = test.EscapeText(ch.ToString());

                // エスケープ対象文字ならば
                if (test.EscapeStringsDic.ContainsKey(ch.ToString()))
                {
                    if (!test.EscapeStringsDic[ch.ToString()].Equals(ans))
                    {
                        Assert.Fail("in -->" + ch.ToString() + " out --> " + ans + "correct -->" + test.EscapeStringsDic[ch.ToString()]);
                    }
                }
                else
                {
                    if (!ch.ToString().Equals(ans))
                    {
                        Assert.Fail("in -->" + ch.ToString() + " out --> " + ans + "correct --> " + ch.ToString());
                    }
                }
            }

            Dictionary<string, string> test_case = new Dictionary<string, string>()
            {
                { "1234567890-^\\qwertyuiop@[asdfghjkl;:]zxcvbnm,./\\", "1234567890-^\\\\qwertyuiop@[asdfghjkl;:]zxcvbnm,./\\\\"},
                { "https://www.premium-tsubu-hero.net/", "https://www.premium-tsubu-hero.net/"},
                { "", ""},
                { "!\"#$%&'()=~|QWERTYUIOP`{ASDFGHJKL+*}ZXCVBNM<>?_", "!\\\"#$%&'()=~|QWERTYUIOP`{ASDFGHJKL+*}ZXCVBNM<>?_"},
                { "今日の天気は晴れ", "今日の天気は晴れ"},
            };

            foreach (var testx in test_case)
            {
                var ansx = test.EscapeText(testx.Key);

                if (!ansx.Equals(testx.Value))
                {
                    Assert.Fail("in -->" + testx.Key + " out --> " + ansx + "correct --> " + testx.Value);
                }

            }
        }
    }
}