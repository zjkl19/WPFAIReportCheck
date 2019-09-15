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
        /// 查找报告中是否存在半角逗号
        /// </summary>
        /// <remarks>
        /// 算法：用正则表达式全文搜索半角逗号，如果查到，则发出警告
        /// </remarks>
        /// 算法主要参与人员：林迪南
        public void FindHalfWidthCommaWarnning()
        {
            FindReplaceOptions options;
            MatchCollection matches;
            var regex = new Regex(@",");
            try
            {
                matches = regex.Matches(_originalWholeText);

                if (matches.Count != 0)
                {
                    foreach (Match m in matches)
                    {
                        reportWarnning.Add(new ReportWarnning(WarnningNumber.FormatProblem, $"正文{m.Index.ToString()}", "请确认此处是否使用半角逗号", true));
                    }

                    options = new FindReplaceOptions
                    {
                        ReplacingCallback = new ReplaceEvaluatorFindAndHighlightWithComment(_doc, "AI校核", "请确认此处是否使用半角逗号"),
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
                _log.Error(ex, $"FindHalfWidthCommaWarnning函数运行出错，错误信息：{ ex.Message.ToString()}");
#endif
            }

        }
    }
}
