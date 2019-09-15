//using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Text;

namespace WPFAIReportCheck.IRepository
{
    public partial interface IAIReportCheck
    {
        void _GenerateResultReport();
        void CheckReport();
        //Table GetOverViewTable();
    }
}
