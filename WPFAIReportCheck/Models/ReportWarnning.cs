using System;
using System.Collections.Generic;
using System.Text;

namespace WPFAIReportCheck.Models
{
    public class ReportWarnning
    {
        public WarnningNumber No;
        public string Name { get; }

        public string Position { get; }

        public string Description { get; }
        public bool HasComment { get; }
        public ReportWarnning(WarnningNumber No, string Position, string Description, bool HasComment = false)
        {
            this.No = No;
            this.Name = Repository.EnumHelper.GetEnumDesc(No).ToString();
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
