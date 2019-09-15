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
        /// 在正文中查找单位错误，并在错误位置建立批注
        /// </summary>
        public void FindUnitError()
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

    }
}

