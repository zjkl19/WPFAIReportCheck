using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using WPFAIReportCheck.IRepository;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        public void CheckReport()
        {
            var config = XDocument.Load(@"Option.config");
            var op = Convert.ToInt32(config.Element("configuration").Element("FindUnitError").Value);
            if (op == 1)
            {
                _FindUnitError();
            }

            
            _FindNotExplainComponentNo();
            _FindSpecificationsError();
            _FindSequenceNumberError();
            FindStrainOrDispError();

            FindDescriptionError();
            FindOtherBridgesWarnning();

            _GenerateResultReport();
            _doc.Save("标出错误或警告的报告.doc");
        }
    }
}
