using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableToJsonlConverter.Conveters;

namespace TableToJsonlConverterTests.Conveters
{
    [TestClass()]
    public class ZkJsonlBaseTests
    {
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
                { "1234567890-^\\qwertyuiop@[asdfghjkl;:]zxcvbnm,./\\", "1234567890-^\\\\qwertyuiop@[asdfghjkl;:]zxcvbnm,.\\/\\\\"},
                { "https://www.premium-tsubu-hero.net/", "https:\\/\\/www.premium-tsubu-hero.net\\/"},
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