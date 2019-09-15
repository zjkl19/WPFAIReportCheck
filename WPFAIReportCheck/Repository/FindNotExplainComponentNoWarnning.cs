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

        public void FindNotExplainComponentNoWarnning()
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
    }
}
