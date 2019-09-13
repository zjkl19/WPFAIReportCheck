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
    public class ProgressBarDataBinding1 
    {

        public string TableTitle { get; set; }

        public Node TableTitleNode { get; set; }

    }

    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        /// <summary>
        /// 查找报告中表格序号的错误
        /// </summary>
        /// <remarks>
        /// 要求：
        /// 表格标题仅支持类似：“表 1-2，表3-2”等，最大支持到“表 999-999”
        /// 算法：
        /// 1、遍历范围：全文表格
        /// 2、在每一张表格中前一个节点的字符串中用正则表达式@"表[\s]{0,1}[1-9][0-9]?[0-9]?[\s]{0,1}-[\s]{0,1}[1-9][0-9]?[0-9]?"查找，将匹配到的第一个结果存到list
        /// 3、将list中的序号从小到大排列，将排序后的list用如下算法进行验证，看是否有错：
        /// 3.1、提取章节号及表序号，第1张表格：一定是x-1
        /// 3.2、第2~n张表格：
        /// 3.3、若章节号同上一张表格，则该表格序号-上一个表序号==1
        /// 3.4、若章节号不同于一张表格，则表序号==1
        /// </remarks>
        /// 算法主要参与人员：林迪南、许智星
        /// TODO：考虑表格标题超过1行的情况
        public void FindTableTitleSequenceNumberError()
        {
            Match m;
            var list = new List<string>();
            var regex = new Regex(@"\b表[\s]{0,1}[1-9][0-9]?[0-9]?[\s]{0,1}-[\s]{0,1}[1-9][0-9]?[0-9]?");
            int previousChapterNum = 0; int previousTableSequenceNum = 0; int currChapterNum = 0; int currTableSequenceNum = 0;
            //var regex = new Regex(@"(?<=\(附[\s]{0,4}页\)\r目录\r)[\s\S]*?(?=\f)"); 
            var chapterNumRegex = new Regex(@"(?<=表)[1-9][0-9]?[0-9]?(?=\-)");
            var tableSequenceNumRegex = new Regex(@"(?<=\-)[1-9][0-9]?[0-9]?");
            try
            {
                list = new List<string>();
                for (int i = 0; i < _doc.GetChildNodes(NodeType.Table, true).Count; i++)
                {
                    try
                    {
                        m = regex.Match(_doc.GetChildNodes(NodeType.Table, true)[i].PreviousSibling.Range.Text);
                        if (m.Success)
                        {
                            list.Add(m.Value.ToString().Replace(" ", "")); ;
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }

                }
                list.Sort();
                for (int i = 0; i < list.Count; i++)
                {
                    currChapterNum = Convert.ToInt32(chapterNumRegex.Match(list[i]).Value);
                    currTableSequenceNum = Convert.ToInt32(tableSequenceNumRegex.Match(list[i]).Value);
                    if (i == 0)
                    {
                        previousChapterNum = Convert.ToInt32(chapterNumRegex.Match(list[i]).Value);
                        previousTableSequenceNum = Convert.ToInt32(tableSequenceNumRegex.Match(list[i]).Value);
                        if (currTableSequenceNum != 1)
                        {
                            reportError.Add(new ReportError(ErrorNumber.Description, $"第1张表", $"第1张表格起始序号应为1", true));
                        }
                    }
                    else
                    {
                        if (currChapterNum == previousChapterNum)
                        {
                            if(currTableSequenceNum- previousTableSequenceNum!=1)
                            {
                                reportError.Add(new ReportError(ErrorNumber.Description, $"第{i+1}张表", $"第{i+1}张表格序号为{currChapterNum}-{currTableSequenceNum}，实际应为{currChapterNum}-{previousTableSequenceNum+1}", true));
                            }
                        }
                        else   //(currChapterNum>previousChapterNum)
                        {
                            if (currTableSequenceNum != 1)
                            {
                                reportError.Add(new ReportError(ErrorNumber.Description, $"第{i+1}张表", $"第{i+1}张表格序号为{currChapterNum}-{currTableSequenceNum}，实际应为{currChapterNum}-1", true));
                            }
                        }
                    }
                    previousChapterNum = currChapterNum;
                    previousTableSequenceNum = currTableSequenceNum;
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                throw ex;

#else
                _log.Error(ex, $"FindTableTitleSequenceNumberError运行出错，错误信息：{ ex.Message.ToString()}");
#endif
            }

        }
    }
}
