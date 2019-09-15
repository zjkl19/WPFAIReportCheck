using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Models;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        /// <summary>
        /// 查找报告中是否存在图片标题单独在1页的现象
        /// </summary>
        /// <remarks>
        /// 要求：
        /// 图片标题命名必须满足正则表达式：图[\s]{0,1}[1-9][0-9]?[0-9]?[\s]{0,1}-[\s]{0,1}[1-9][0-9]?[0-9]?[\s]{1}(.*)
        /// 该正则表达式表示表的索引范围：“图 1-1” 至 “图 999-999”
        /// 算法：
        /// 1、根据要求的正则表达式，在渲染后的word中匹配该字符串
        /// 2、遍历所有页，若匹配到的字符串在页的第1行（并且该行上方没有图片），将该字符串存储起来，记为{s}，并添加警告reportWarnning：图片标题不能单独放在1页
        /// 3、_doc从第2节开始，在每个Paragraph中，查找{s}，若找到，则添加common，退出循环，算法结束。
        /// </remarks>
        /// 算法主要参与人员：林迪南、许智星
        /// TODO：考虑图片标题超过1行的情况
        public void FindPictureTitleInSinglePageWarnning()
        {
            MatchCollection matches;
            var regex = new Regex(@"图[\s]{0,1}[1-9][0-9]?[0-9]?[\s]{0,1}-[\s]{0,1}[1-9][0-9]?[0-9]?[\s]{1}(.*)");
            string matchTitle = string.Empty;
            int i = 0;

            var pageSetupBuilder = new DocumentBuilder(_doc);    //获取页面设置
            var pageSetup = _doc.Sections[0].PageSetup;

            try
            {
                for (i = 0; i < _layoutDoc.Pages.Count; i++)
                {
                    if (_layoutDoc.Pages[i].Columns[0].Lines.Count == 0)    //有些页全是图片或表格，没有文字
                    {
                        continue;
                    }
                    
                    matches = regex.Matches(_layoutDoc.Pages[i].Columns[0].Lines.First.Text);
                    
                    var firstLineRectangle = _layoutDoc.Pages[i].Columns[0].Lines.First.Rectangle;
                    //触发警告条件：第1行找到关键字，并且第1行在页面顶端
                    //第1行在页面顶端判定条件：firstLineRectangle.Top < pageSetup.TopMargin + 20
                    if (matches.Count > 0 && firstLineRectangle.Top < pageSetup.TopMargin + 20)
                    {
                        matchTitle = matches[0].Value.ToString();
                        reportWarnning.Add(new ReportWarnning(WarnningNumber.FormatProblem, $"第{i + 1}页第1行", $"{matchTitle}位于第{i + 1}页第1行", true));

                        Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Now);
                        comment.Paragraphs.Add(new Paragraph(_doc));
                        comment.FirstParagraph.Runs.Add(new Run(_doc, $"图片标题单独位于某1页"));
                        DocumentBuilder builder = new DocumentBuilder(_doc);
                        bool flag = true;
                        for (int j = 1; j < _doc.GetChildNodes(NodeType.Section, true).Count; j++)    //从第2节开始搜索
                        {
                            //TODO：找不到
                            for (int k = 0; k < _doc.Sections[j].GetChildNodes(NodeType.Paragraph, true).Count; k++)
                            {
                                if (_doc.Sections[j].GetChildNodes(NodeType.Paragraph, true)[k].Range.Text.Contains(matchTitle.Replace("¶", "").Replace("\r", "")))
                                {
                                    builder.MoveTo(_doc.Sections[j].GetChildNodes(NodeType.Paragraph, true)[k] as Paragraph);
                                    builder.CurrentParagraph.AppendChild(comment);
                                    flag = false;
                                    break;
                                }
                            }
                            if (flag == false)
                            {
                                break;
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                throw ex;
#else
                _log.Error(ex, $"FindPictureTitleInSinglePageWarnning在渲染后的文档第{i}页出错，错误信息：{ ex.Message.ToString()}");
#endif
            }
        }
    }
}
