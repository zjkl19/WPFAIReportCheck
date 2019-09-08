using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Tables;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Models;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        /// <summary>
        /// 查找报告中目录页码和正文页码对应不上的错误
        /// </summary>
        /// <remarks>
        /// 要求：
        /// 1、目录页必须在第一节
        /// 2、目录与第一章必须有分节符\f
        /// 算法：
        /// 1、根据正则表达式：以"(附 页)\r目录\r"开头，以"\f"结尾匹配字符串
        /// 2、以"\r"分割字符串为{s1},{s2},...,{si},...,{sn}
        /// 3、每个被分割的字符串{si}再分别用正则表达式Regex(@"(?<=\t)[1-9]\d*")，Regex(@"[\s\S]*(?=\t)")分割成2个字符串，
        /// 第1个字符串为每1小节页码{1}，第2个字符串为每1小节正文{2}
        /// 4、在渲染后的doc的第{1}页查找{2}，若找到，则继续对{si+1}字符串执行3步骤，直到i==n
        /// </remarks>
        /// 算法主要参与人员：林迪南
        public void FindPageContextError()
        {
            MatchCollection matches;
            //var regex = new Regex(@"(?<=\(附[\s]{0,4}页\)\\r目录\\r)[\s\S]*?(?=\\f)");    //https://regex101.com/使用   
            var regex = new Regex(@"(?<=\(附[\s]{0,4}页\)\r目录\r)[\s\S]*?(?=\f)");    //在第1节中查找目录用的正则表达式字符串
            var regexPageNumber=new Regex(@"(?<=\t)[1-9]\d*");    //每1小节查找页码使用的正则表达式字符串
            var regexContent = new Regex(@"[\s\S]*(?=\t)");    //每1小节查找正文使用的正则表达式字符串
            try
            {
                matches = regex.Matches(_doc.FirstSection.Range.Text);
                if (matches.Count != 0)
                {
                    _log.Debug(matches[0].Value);
                    var list = new List<string>(matches[0].Value.Split(new[] { "\r" }, StringSplitOptions.RemoveEmptyEntries));//返回值不包括含有空字符串的数组元素
                    for (int i = 0; i < list.Count; i++)
                    {
                        _log.Debug(list[i]);

                        regex = regexPageNumber;
                        matches = regex.Matches(list[i]);
                        _log.Debug(matches[0].Value);
                        int pageNumber = Convert.ToInt32(matches[0].Value.Replace(@"\r", ""));

                        regex = regexContent;
                        matches = regex.Matches(list[i]);
                        _log.Debug(matches[0].Value);
                        string content = matches[0].Value;
                        if (pageNumber > _layoutDoc.Pages.Count)
                        {
                            reportError.Add(new ReportError(ErrorNumber.Description, $"目录", $"正文只有{_layoutDoc.Pages.Count}页，但页码索引却有{pageNumber}", true));

                            Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Now);
                            comment.Paragraphs.Add(new Paragraph(_doc));
                            comment.FirstParagraph.Runs.Add(new Run(_doc, $"正文只有{_layoutDoc.Pages.Count}页，但页码索引却有{pageNumber}"));
                            DocumentBuilder builder = new DocumentBuilder(_doc);
                            //TODO：_doc.GetChildNodes(NodeType.Table, true)[i] as Table).PreviousSibling.PreviousSibling 精确定位
                            for (int j = 0; j < _doc.FirstSection.GetChildNodes(NodeType.Paragraph, true).Count; j++)
                            {
                                if (_doc.FirstSection.GetChildNodes(NodeType.Paragraph, true)[j].Range.Text.Contains("(附 页)"))
                                {
                                    builder.MoveTo((_doc.GetChildNodes(NodeType.Paragraph, true)[j] as Paragraph));
                                    builder.CurrentParagraph.AppendChild(comment);
                                    break;
                                }
                            }

                       }
                        else
                        {
                            if(!_layoutDoc.Pages[pageNumber - 1].Text.Contains(content))
                            {
                                reportError.Add(new ReportError(ErrorNumber.Description, $"正文第{pageNumber}页或目录", $"正文第{ pageNumber }页未找到{ content }", true));

                                Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Now);
                                comment.Paragraphs.Add(new Paragraph(_doc));
                                comment.FirstParagraph.Runs.Add(new Run(_doc, $"正文第{ pageNumber }页未找到{ content }"));
                                DocumentBuilder builder = new DocumentBuilder(_doc);
                                for (int j = 0; j < _doc.FirstSection.GetChildNodes(NodeType.Paragraph, true).Count; j++)
                                {
                                    if (_doc.FirstSection.GetChildNodes(NodeType.Paragraph, true)[j].Range.Text.Contains("(附 页)"))
                                    {
                                        builder.MoveTo((_doc.GetChildNodes(NodeType.Paragraph, true)[j] as Paragraph));
                                        builder.CurrentParagraph.AppendChild(comment);
                                        break;
                                    }
                                }
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
                _log.Error(ex,$"目录上下文校核出错，错误信息：{ ex.Message.ToString()}");
#endif
            }
        }
    }
}
