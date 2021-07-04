using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddIn1
{
    public static class WordHelper
    {
        /// <summary>
        /// 查找所有
        /// </summary>
        /// <param name="range"></param>
        /// <param name="findText">查找内容</param>
        public static List<Range> FindAll(this Range range, string findText)
        {
            int start = range.Start;
            int end = range.End;

            List<Range> ranges = new List<Range>();

            range.Find.Execute(FindText: findText, MatchCase: true);
            while (range.Find.Found)
            {
                //搜索会改变range，这里做了一个超出范围的判断
                if (range.Start > end)
                    break;

                ranges.Add(range.Document.Range(range.Start, range.End));
                range.Find.Execute(FindText: findText, MatchCase: true);
            }

            //对原来的range还原
            range.SetRange(start, end);
            return ranges;
        }
        /// <summary>
        /// 根据模式，找到所有匹配的位置
        /// </summary>
        /// <param name="range">选中部分</param>
        /// <param name="pattern">模式</param>
        /// <returns>匹配列表</returns>
        public static List<Range> SearchRangeInPattern(this Range range, string pattern)
        {
            var ret = new List<Range> { };

            var content = range.Text;
            var doc = range.Document;

            //获取实际的字符位置
            var locationList = GetRangeLocation(range);

            var mc = System.Text.RegularExpressions.Regex.Matches(content, pattern);

            if (mc.Count > 0)
            {
                foreach (System.Text.RegularExpressions.Match m in mc)
                {
                    var searchStart = m.Index;
                    var searchEnd = m.Index + m.Value.Length;

                    //将text位置转换为range位置
                    var realStart = locationList[searchStart];
                    var realEnd = locationList[searchEnd];

                    //获取匹配的range位置
                    var itemRange = doc.Range(realStart, realEnd);

                    ret.Add(itemRange);
                }
            }

            return ret;

        }
        /// <summary>
        /// 从range中获取每一个字符的实际位置
        /// </summary>
        /// <param name="range">选中部分</param>
        /// <returns>位置列表</returns>
        public static List<int> GetRangeLocation(Range range)
        {
            var ret = new List<int> { };

            foreach (Range c in range.Characters)
            {
                ret.Add(c.Start);
            }

            return ret;
        }
    }
}
