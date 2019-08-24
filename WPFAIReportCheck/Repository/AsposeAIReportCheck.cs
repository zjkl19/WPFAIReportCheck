using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Models;
//using Microsoft.Office.Interop.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
//using log4net;
//using log4net.Repository;
//using log4net.Config;
using System.Xml.Linq;
using NLog;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        public List<ReportError> reportError = new List<ReportError>();
        public List<ReportWarnning> reportWarnning = new List<ReportWarnning>();
        public Document _doc;
        //public ILog _log;
        public ILogger _log;
        public XDocument _config;
        readonly string _originalWholeText;
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="doc">doc文件，默认和程序在同一个目录</param>
        public AsposeAIReportCheck(string doc, ILogger log,XDocument config)
        {
            _doc = new Document(doc);
            _originalWholeText = _doc.Range.Text;

            //ILoggerRepository repository = LogManager.CreateRepository("WPFAIReportCheck");
            // 默认简单配置，输出至控制台
            //BasicConfigurator.Configure(repository);
            // _log = log4net.LogManager.GetLogger(repository.Name, "WPFAIReportCheckLog4net");
            _log = log;
            _config = config;
        }

        /// <summary>
        /// 在正文中查找单位错误，并在错误位置建立批注
        /// </summary>
        public void _FindUnitError()
        {
            FindReplaceOptions options;
            MatchCollection matches;
            var regex = new Regex(@"([0-9]Km/h)");
            try
            {
                matches = regex.Matches(_originalWholeText);

                if (matches.Count != 0)
                {
                    foreach (Match m in matches)
                    {
                        reportError.Add(new ReportError(ErrorNumber.CMA, "正文" + m.Index.ToString(), "应为km/h", true));
                    }

                    options = new FindReplaceOptions
                    {
                        ReplacingCallback = new ReplaceEvaluatorFindAndHighlightWithComment(_doc, "AI校核", "应为km/h"),
                        //new ReplaceEvaluatorFindAndHighlight(_doc),
                        Direction = FindReplaceDirection.Forward
                    };
                    _doc.Range.Replace(regex, "", options);

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void _FindSpecificationsError()
        {
            string[] Specifications;
            //规范
            try
            {
                var file = new System.IO.StreamReader("规范.txt", Encoding.Default);
                string line;
                string combineString = string.Empty;
                while ((line = file.ReadLine()) != null)
                {
                    combineString += line + '\r';
                }
                Specifications = combineString.Split('\r');
                file.Close();
            }
            catch (Exception)
            {
                Specifications = new string[]
                {
                    "《城市桥梁设计规范》（CJJ 11-2011）",
                    "《混凝土结构现场检测技术标准》（GB/T 50784-2013）",
                    "《公路桥梁荷载试验规程》（JTG/T J21-01-2015）",
                    "《城市桥梁检测与评定技术规范》（CJJ/T 233-2015）",
                    "《城市桥梁养护技术标准》（CJJ 99-2017）",
                };
            }

            double similarity;    //相似度
                                  //获取word文档中的第一个表格

            //Table table0=null;
            //Cell cell=null;
            //FindReplaceOptions options;
            //string repStr=string.Empty;
            double similarityLBound = 0.85;
            double similarityUBound = 1.00;

            NodeCollection allTables = _doc.GetChildNodes(NodeType.Table, true);
            for (int i = 0; i < allTables.Count; i++)
            {
                Table table0 = _doc.GetChildNodes(NodeType.Table, true)[i] as Table;
                if ((table0.Rows[0].Cells[0].GetText().IndexOf("委托单位") >= 0))
                {
                    Cell cell = table0.Rows[4].Cells[1];
                    string[] splitArray = cell.GetText().Split('\r');    //用GetText()的方法来获取cell中的值
                    foreach (var s in splitArray)
                    {
                        string repStr = s;
                        repStr.Replace("\a", ""); repStr.Replace("\r", "");
                        repStr.Replace("(", "（"); repStr.Replace(")", "）");
                        var s1 = Regex.Replace(repStr, @"(.+)《", "《");    //替换"《"之前的内容为""
                        foreach (var sp in Specifications)
                        {
                            similarity = Levenshtein(@s1, @sp);
                            if (similarity > similarityLBound && similarity < similarityUBound)
                            {
                                var regex = new Regex(@s);
                                reportError.Add(new ReportError(ErrorNumber.Description, "汇总表格中主要检测检验依据", "应为" + sp));
                                FindReplaceOptions options = new FindReplaceOptions
                                {
                                    ReplacingCallback = new ReplaceEvaluatorFindAndHighlightWithComment(_doc, "AI校核", "应为" + sp),
                                    Direction = FindReplaceDirection.Forward
                                };
                                _doc.Range.Replace(regex, "", options);
                                break;
                            }
                        }
                    };
                    break;
                }
            }
            //var table0 = ai.GetOverViewTable();
        }

        public void _FindNotExplainComponentNo()
        {
            MatchCollection matches;
            var regex = new Regex(@"(构件编号说明)");
            try
            {
                matches = regex.Matches(_originalWholeText);
                if (matches.Count == 0)
                {
                    reportWarnning.Add(new ReportWarnning(WarnningNumber.NotClearInfo, "/", "构件编号未进行说明"));
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void _FindSequenceNumberError()
        {
            NodeCollection allTables = _doc.GetChildNodes(NodeType.Table, true);
            for (int i = 0; i < allTables.Count; i++)
            {
                Table table0 = _doc.GetChildNodes(NodeType.Table, true)[i] as Table;
                if (table0.Rows[0].Cells[0].GetText().IndexOf("序号") >= 0)    //含有序号的表格
                {
                    int sn = 1;
                    for (int j = 0; j < table0.IndexOf(table0.LastRow); j++)
                    {
                        var sn1 = table0.Rows[j + 1].Cells[0].GetText().Replace("\a", "").Replace("\r", "");
                        try
                        {
                            //TODO:增加转换失败的测试
                            if (Convert.ToInt32(sn1) != sn)    //转换可能会失败（如序号中含有中文）
                            {
                                reportError.Add(new ReportError(ErrorNumber.Description, $"第{i + 1}张表格", "序号应连贯，从小到大", true));

                                Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Today);
                                comment.Paragraphs.Add(new Paragraph(_doc));
                                comment.FirstParagraph.Runs.Add(new Run(_doc, "序号应连贯，从小到大"));
                                DocumentBuilder builder = new DocumentBuilder(_doc);
                                builder.MoveTo(table0.Rows[j + 1].Cells[0].FirstParagraph);
                                builder.CurrentParagraph.AppendChild(comment);
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            throw ex;

#else
                            //TODO：增加错误定位信息
                            _log.Error(ex, $"FindSequenceNumberError函数运行出错，错误信息：{ ex.Message.ToString()}");
#endif
                        }

                        sn++;
                    }
                }
            }
        }
        
        public void _GenerateResultReport()
        {
            var doc = new Document();

            // DocumentBuilder provides members to easily add content to a document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            if (reportError.Count != 0)
            {
                builder.Writeln("错误列表");
                int i = 1;    //计数
                builder.StartTable();
                builder.InsertCell(); builder.Write("序号");
                builder.InsertCell(); builder.Write("编号");
                builder.InsertCell(); builder.Write("名称");
                builder.InsertCell(); builder.Write("位置");
                builder.InsertCell(); builder.Write("说明");
                builder.InsertCell(); builder.Write("批注");
                builder.EndRow();

                foreach (var e in reportError)
                {
                    builder.InsertCell(); builder.Write(i.ToString());
                    builder.InsertCell(); builder.Write(((int)e.No).ToString());
                    builder.InsertCell(); builder.Write(e.Name);
                    builder.InsertCell(); builder.Write(e.Position);
                    builder.InsertCell(); builder.Write(e.Description);
                    builder.InsertCell(); builder.Write(e.HasComment.ToString());
                    builder.EndRow();
                    i += 1;
                }
                builder.EndTable();
            }

            if (reportWarnning.Count != 0)
            {
                builder.Writeln("警告列表");
                int i = 1;    //计数
                builder.StartTable();
                builder.InsertCell(); builder.Write("序号");
                builder.InsertCell(); builder.Write("编号");
                builder.InsertCell(); builder.Write("名称");
                builder.InsertCell(); builder.Write("位置");
                builder.InsertCell(); builder.Write("说明");
                builder.InsertCell(); builder.Write("批注");
                builder.EndRow();

                foreach (var w in reportWarnning)
                {
                    builder.InsertCell(); builder.Write(i.ToString());
                    builder.InsertCell(); builder.Write(((int)w.No).ToString());
                    builder.InsertCell(); builder.Write(w.Name);
                    builder.InsertCell(); builder.Write(w.Position);
                    builder.InsertCell(); builder.Write(w.Description);
                    builder.InsertCell(); builder.Write(w.HasComment.ToString());
                    builder.EndRow();
                    i += 1;
                }
                builder.EndTable();
            }

            // Save the document in DOCX format. The format to save as is inferred from the extension of the file name.
            // Aspose.Words supports saving any document in many more formats.
            doc.Save("校核结果.docx");
        }

        /// <summary>
        /// 字符串相似度计算
        /// </summary>
        /// <param name="str1">字符串1</param>
        /// <param name="str2">字符串2</param>
        /// <returns>两个字符串的相似度</returns>
        public static double Levenshtein(string str1, string str2)
        {
            //计算两个字符串的长度。  
            int len1 = str1.Length;
            int len2 = str2.Length;
            //建立上面说的数组，比字符长度大一个空间  
            int[,] dif = new int[len1 + 1, len2 + 1];
            //赋初值，步骤B。  
            for (int a = 0; a <= len1; a++)
            {
                dif[a, 0] = a;
            }
            for (int a = 0; a <= len2; a++)
            {
                dif[0, a] = a;
            }
            //计算两个字符是否一样，计算左上的值  
            int temp;
            for (int i = 1; i <= len1; i++)
            {
                for (int j = 1; j <= len2; j++)
                {
                    if (str1[i - 1] == str2[j - 1])
                    {
                        temp = 0;
                    }
                    else
                    {
                        temp = 1;
                    }
                    //取三个值中最小的  
                    dif[i, j] = Math.Min(Math.Min(dif[i - 1, j - 1] + temp, dif[i, j - 1] + 1), dif[i - 1, j] + 1);
                }
            }
            //Console.WriteLine("字符串\"" + str1 + "\"与\"" + str2 + "\"的比较");
            //取数组右下角的值，同样不同位置代表不同字符串的比较  
            //Console.WriteLine("差异步骤：" + dif[len1, len2]);
            //计算相似度  
            double similarity = 1 - (double)dif[len1, len2] / Math.Max(str1.Length, str2.Length);
            //Console.WriteLine("相似度：" + similarity + " 越接近1越相似");
            return similarity;
        }
        //public static void Run()
        //{
        //    // ExStart:FindAndHighlight
        //    // The path to the documents directory.

        //    string fileName = "TestFile.doc";

        //    Document doc = new Document(fileName);

        //    FindReplaceOptions options = new FindReplaceOptions();
        //    options.ReplacingCallback = new ReplaceEvaluatorFindAndHighlight();
        //    options.Direction = FindReplaceDirection.Forward;

        //    // We want the "your document" phrase to be highlighted.
        //    Regex regex = new Regex("your document", RegexOptions.IgnoreCase);
        //    doc.Range.Replace(regex, "", options);

        //    // Save the output document.
        //    doc.Save("TestFile.doc");
        //    // ExEnd:FindAndHighlight
        //}
        // ExStart:ReplaceEvaluatorFindAndHighlight

        private class ReplaceEvaluatorFindAndHighlightWithComment : IReplacingCallback
        {
            private Document _doc;
            private string _initialText;
            private string _commentText;
            /// <summary>
            /// 正则表达式替换文字加高亮和批注
            /// </summary>
            /// <param name="doc">Aspose文档</param>
            /// <param name="initialText"></param>
            /// <param name="commentText">批注文档</param>
            public ReplaceEvaluatorFindAndHighlightWithComment(Document doc, string initialText, string commentText)
            {
                _doc = doc;
                _initialText = initialText;
                _commentText = commentText;
            }
            /// <summary>
            /// This method is called by the Aspose.Words find and replace engine for each match.
            /// This method highlights the match string, even if it spans multiple runs.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                // This is a Run node that contains either the beginning or the complete match.
                Node currentNode = e.MatchNode;

                // The first (and may be the only) run can contain text before the match, 
                // In this case it is necessary to split the run.
                if (e.MatchOffset > 0)
                    currentNode = SplitRun((Run)currentNode, e.MatchOffset);

                // This array is used to store all nodes of the match for further highlighting.
                ArrayList runs = new ArrayList();

                // Find all runs that contain parts of the match string.
                int remainingLength = e.Match.Value.Length;
                while (
                    (remainingLength > 0) &&
                    (currentNode != null) &&
                    (currentNode.GetText().Length <= remainingLength))
                {
                    runs.Add(currentNode);
                    remainingLength -= currentNode.GetText().Length;

                    // Select the next Run node. 
                    // Have to loop because there could be other nodes such as BookmarkStart etc.
                    do
                    {
                        currentNode = currentNode.NextSibling;
                    }
                    while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
                }

                // Split the last run that contains the match if there is any text left.
                if ((currentNode != null) && (remainingLength > 0))
                {
                    SplitRun((Run)currentNode, remainingLength);
                    runs.Add(currentNode);
                }

                // Now highlight all runs in the sequence.
                //foreach (Run run in runs)
                //    run.Font.HighlightColor = System.Drawing.Color.Red;

                Comment comment = new Comment(_doc, "AI", _initialText, DateTime.Today);
                comment.Paragraphs.Add(new Paragraph(_doc));
                comment.FirstParagraph.Runs.Add(new Run(_doc, _commentText));

                CommentRangeStart commentRangeStart = new CommentRangeStart(_doc, comment.Id);
                CommentRangeEnd commentRangeEnd = new CommentRangeEnd(_doc, comment.Id);

                //run1.ParentNode.InsertAfter(commentRangeStart, run1);
                //run3.ParentNode.InsertAfter(commentRangeEnd, run3);
                //commentRangeEnd.ParentNode.InsertAfter(comment, commentRangeEnd);

                foreach (Run run in runs)
                {
                    run.Font.HighlightColor = System.Drawing.Color.Red;
                    run.ParentNode.InsertAfter(commentRangeStart, run);
                    run.ParentNode.InsertAfter(commentRangeEnd, run);
                    commentRangeEnd.ParentNode.InsertAfter(comment, commentRangeEnd);

                }

                //Comment comment = new Comment(doc, "Awais Hafeez", "AH", DateTime.Today);
                //comment.Paragraphs.Add(new Paragraph(doc));
                //comment.FirstParagraph.Runs.Add(new Run(doc, "Comment text."));

                //CommentRangeStart commentRangeStart = new CommentRangeStart(doc, comment.Id);
                //CommentRangeEnd commentRangeEnd = new CommentRangeEnd(doc, comment.Id);

                //run1.ParentNode.InsertAfter(commentRangeStart, run1);
                //run3.ParentNode.InsertAfter(commentRangeEnd, run3);
                //commentRangeEnd.ParentNode.InsertAfter(comment, commentRangeEnd);

                // Signal to the replace engine to do nothing because we have already done all what we wanted.
                return ReplaceAction.Skip;
            }
        }

        private class ReplaceEvaluatorFindAndHighlight : IReplacingCallback
        {
            private Document _doc;
            public ReplaceEvaluatorFindAndHighlight(Document doc)
            {
                _doc = doc;
            }
            /// <summary>
            /// This method is called by the Aspose.Words find and replace engine for each match.
            /// This method highlights the match string, even if it spans multiple runs.
            /// </summary>
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                // This is a Run node that contains either the beginning or the complete match.
                Node currentNode = e.MatchNode;

                // The first (and may be the only) run can contain text before the match, 
                // In this case it is necessary to split the run.
                if (e.MatchOffset > 0)
                    currentNode = SplitRun((Run)currentNode, e.MatchOffset);

                // This array is used to store all nodes of the match for further highlighting.
                ArrayList runs = new ArrayList();

                // Find all runs that contain parts of the match string.
                int remainingLength = e.Match.Value.Length;
                while (
                    (remainingLength > 0) &&
                    (currentNode != null) &&
                    (currentNode.GetText().Length <= remainingLength))
                {
                    runs.Add(currentNode);
                    remainingLength -= currentNode.GetText().Length;

                    // Select the next Run node. 
                    // Have to loop because there could be other nodes such as BookmarkStart etc.
                    do
                    {
                        currentNode = currentNode.NextSibling;
                    }
                    while ((currentNode != null) && (currentNode.NodeType != NodeType.Run));
                }

                // Split the last run that contains the match if there is any text left.
                if ((currentNode != null) && (remainingLength > 0))
                {
                    SplitRun((Run)currentNode, remainingLength);
                    runs.Add(currentNode);
                }

                // Now highlight all runs in the sequence.
                //foreach (Run run in runs)
                //    run.Font.HighlightColor = System.Drawing.Color.Red;

                Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Today);
                comment.Paragraphs.Add(new Paragraph(_doc));
                comment.FirstParagraph.Runs.Add(new Run(_doc, "单位错误"));

                CommentRangeStart commentRangeStart = new CommentRangeStart(_doc, comment.Id);
                CommentRangeEnd commentRangeEnd = new CommentRangeEnd(_doc, comment.Id);

                //run1.ParentNode.InsertAfter(commentRangeStart, run1);
                //run3.ParentNode.InsertAfter(commentRangeEnd, run3);
                //commentRangeEnd.ParentNode.InsertAfter(comment, commentRangeEnd);

                foreach (Run run in runs)
                {
                    run.Font.HighlightColor = System.Drawing.Color.Red;
                    run.ParentNode.InsertAfter(commentRangeStart, run);
                    run.ParentNode.InsertAfter(commentRangeEnd, run);
                    commentRangeEnd.ParentNode.InsertAfter(comment, commentRangeEnd);

                }

                //Comment comment = new Comment(doc, "Awais Hafeez", "AH", DateTime.Today);
                //comment.Paragraphs.Add(new Paragraph(doc));
                //comment.FirstParagraph.Runs.Add(new Run(doc, "Comment text."));

                //CommentRangeStart commentRangeStart = new CommentRangeStart(doc, comment.Id);
                //CommentRangeEnd commentRangeEnd = new CommentRangeEnd(doc, comment.Id);

                //run1.ParentNode.InsertAfter(commentRangeStart, run1);
                //run3.ParentNode.InsertAfter(commentRangeEnd, run3);
                //commentRangeEnd.ParentNode.InsertAfter(comment, commentRangeEnd);

                // Signal to the replace engine to do nothing because we have already done all what we wanted.
                return ReplaceAction.Skip;
            }
        }
        // ExEnd:ReplaceEvaluatorFindAndHighlight
        // ExStart:SplitRun
        /// <summary>
        /// Splits text of the specified run into two runs.
        /// Inserts the new run just after the specified run.
        /// </summary>
        private static Run SplitRun(Run run, int position)
        {
            Run afterRun = (Run)run.Clone(true);
            afterRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);
            return afterRun;
        }
        // ExEnd:SplitRun
    }
}
