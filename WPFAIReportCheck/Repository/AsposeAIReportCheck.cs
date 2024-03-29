﻿using Aspose.Words;
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
        public List<ReportInfo> reportInfo = new List<ReportInfo>();
        public Document _doc;

        private Document _originalDoc;
        private RenderedDocument _layoutDoc;

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

            //_originalDoc = new Document(doc);
            _originalWholeText = _doc.Range.Text;

            //TODO：重构以下5行代码
            _doc.Unprotect();    //解除保护
            _doc.UpdateFields();
            _originalDoc = _doc.Clone();
            _doc.UnlinkFields();    //章节序号等尚未解除链接
            _layoutDoc = new RenderedDocument(_originalDoc);

            //ILoggerRepository repository = LogManager.CreateRepository("WPFAIReportCheck");
            // 默认简单配置，输出至控制台
            //BasicConfigurator.Configure(repository);
            // _log = log4net.LogManager.GetLogger(repository.Name, "WPFAIReportCheckLog4net");
            _log = log;
            _config = config;
        }
      
        public void GenerateResultReport()
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
