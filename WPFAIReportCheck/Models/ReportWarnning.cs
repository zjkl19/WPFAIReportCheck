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
        public bool HasComment { get; }
        public ReportWarnning(WarnningNumber No, string Position, string Description, bool HasComment = false)
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
            this.HasComment = HasComment;
        }
    }
    public enum WarnningNumber
    {
        NotClearInfo = 1,
    }
}
