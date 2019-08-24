using System;
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
        /// 查找报告中名词描述错误
        /// </summary>
        /// <remarks>
        /// 算法：用正则表达式全文搜索“CNAS检查机构”关键字，如果查到，表示有误
        /// TODO：增加其它名词描述错误
        /// </remarks>
        /// 算法主要参与人员：林迪南
        public void FindDescriptionError()
        {
            FindReplaceOptions options;
            MatchCollection matches;
            var regex = new Regex(@"CNAS检查机构");
            try
            {
                matches = regex.Matches(_originalWholeText);

                if (matches.Count != 0)
                {
                    foreach (Match m in matches)
                    {
                        reportError.Add(new ReportError(ErrorNumber.Description, "正文" + m.Index.ToString(), "应为\"CNAS检验机构\"", true));
                    }

                    options = new FindReplaceOptions
                    {
                        ReplacingCallback = new ReplaceEvaluatorFindAndHighlightWithComment(_doc, "AI校核", "应为\"CNAS检验机构\""),
                        Direction = FindReplaceDirection.Forward
                    };
                    _doc.Range.Replace(regex, "", options);

                }
            }
            catch (Exception ex)
            {
#if DEBUG
                throw ex;

#else
                            _log.Error(ex, $"FindDescriptionError函数运行出错，错误信息：{ ex.Message.ToString()}");
#endif
            }
 
        }
    }
}
