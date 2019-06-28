//using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Text;

namespace WPFAIReportCheck.IRepository
{
    public interface IAIReportCheck
    {
        void _FindUnitError();
        void _FindSpecificationsError();
        void _FindNotExplainComponentNo();

        void _FindSequenceNumberError();

        void _GenerateResultReport();
        void CheckReport();
        //Table GetOverViewTable();
    }
}
