using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Models;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        /// <summary>
        /// 查找报告中的重复桥梁，若查找到疑似重复的桥梁，则发出警告
        /// </summary>
        /// <remarks>
        /// 算法：
        /// 步骤：
        /// 1、查找报告使用的桥梁。算法：遍历表格，查找到汇总表格（1行1列含有"委托单位"四个字），取3行2列的内容作为报告的桥梁名称
        /// 2、
        /// 匹配到的正则表达式如果不是“报告桥梁名称”的一部分，则发出警告
        /// </remarks>
        /// 算法主要参与人员：林迪南，陈思远
        public void FindOtherBridgesWarnning()
        {
            var bridgeName = string.Empty;
            NodeCollection allTables = _doc.GetChildNodes(NodeType.Table, true);
            for (int i = 0; i < allTables.Count; i++)
            {
                Table table0 = _doc.GetChildNodes(NodeType.Table, true)[i] as Table;
                if ((table0.Rows[0].Cells[0].GetText().IndexOf("委托单位") >= 0))
                {
                    Cell cell = table0.Rows[2].Cells[1];
                    bridgeName = cell.GetText().Replace("\a", "").Replace("\r", "");    //用GetText()的方法来获取cell中的值

                    break;
                }
            }

            FindReplaceOptions options;
            MatchCollection matches;
            //参考：https://www.runoob.com/regexp/regexp-metachar.html
            //"."匹配除换行符（\n、\r）之外的任何单个字符。要匹配包括 '\n' 在内的任何字符，请使用像"(.|\n)"的模式
            //在线正则表达式测试https://regex101.com/
            //1、县或市开头（但不进行匹配）
            //2、1~20个任意字符（\n除外）并且不以"该"结尾，懒惰匹配
            //3、"大桥"，"小桥"或"桥"结尾，但不以"桥梁"，"桥面"结尾

            //20190825:(?<=[县|市])(.{1,20}[^该])?[大|小]?桥(?![梁|面])
            var regex = new Regex(@"(?=[县|市])(.[^，。》]{1,20}[^该])?[大|小]?桥(?![梁|面])");


            try
            {
                var matchResult = new List<string>();    //去重结果
                var matchResultIndex = new List<int>();    //去重结果位置
                matches = regex.Matches(_originalWholeText);

                if (matches.Count != 0)
                {
                    foreach (Match m in matches)
                    {
#if DEBUG
                        _log.Debug(m.Value.ToString());
#endif
                        if(!matchResult.Contains(m.Value.ToString()))
                        {
                            matchResult.Add(m.Value.ToString());
                            matchResultIndex.Add(m.Index);
                        }
                        
                    }
                }
                for(int i=0;i<matchResult.Count;i++)
                {
                    //也可以用IndexOf
                    if (!bridgeName.Contains(matchResult[i]))    //查看匹配到的正则表达式是不是桥梁名称的一部分
                    {
                        reportWarnning.Add(new ReportWarnning(WarnningNumber.NotClearInfo, "正文" + matchResultIndex[i], $"报告中疑似出现其它桥梁：{matchResult[i]}", true));

                        //TODO：效率低，有重复，要改
                        options = new FindReplaceOptions
                        {
                            ReplacingCallback = new ReplaceEvaluatorFindAndHighlightWithComment(_doc, "AI校核", $"报告中疑似出现其它桥梁：{matchResult[i]}"),
                            Direction = FindReplaceDirection.Forward
                        };
                        _doc.Range.Replace(matchResult[i], "", options);
                    }
                }
            }
            catch (Exception ex)
            {
#if DEBUG

                throw ex;

#else
                            _log.Error(ex, $"FindDuplicateBridgesWarnning函数运行出错，错误信息：{ ex.Message.ToString()}");
#endif
            }

        }
    }
}
