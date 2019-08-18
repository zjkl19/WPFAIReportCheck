using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WPFAIReportCheck.IRepository;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        public void CheckReport()
        {
            _FindUnitError();
            _FindNotExplainComponentNo();
            _FindSpecificationsError();
            _FindSequenceNumberError();
            FindStrainOrDispError();

            FindDescriptionError();

            _GenerateResultReport();
            _doc.Save("标出错误或警告的报告.doc");
        }
    }
}
