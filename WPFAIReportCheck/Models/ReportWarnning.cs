using System;
using System.Collections.Generic;
using System.Text;

namespace WPFAIReportCheck.Models
{
    class ReportWarnning
    {
        public WarnningNumber No;
        public string Name { get; }

        public string Position { get; }

        public string Description { get; }
        public ReportWarnning(WarnningNumber No, string Position, string Description)
        {
            this.No = No;
            if (No == WarnningNumber.NotClearInfo)
            {
                Name = "信息不明确";
            }
            else
            {
                Name = "待定";
            }
            this.Position = Position;
            this.Description = Description;
        }
    }
    public enum WarnningNumber
    {
        NotClearInfo = 1,
    }
}
