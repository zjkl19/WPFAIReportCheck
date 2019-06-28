using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPFAIReportCheck.Models
{
    public class ReportInfo
    {
        public InfoNumber No;
        public string Name { get; }

        public string Position { get; }

        public string Description { get; }
        public bool HasComment { get; }
        public ReportInfo(InfoNumber No, string Position, string Description, bool HasComment = false)
        {
            this.No = No;
            this.Name = Repository.EnumHelper.GetEnumDesc(No).ToString();
            this.Position = Position;
            this.Description = Description;
            this.HasComment = HasComment;
        }
    }
    public enum InfoNumber
    {
        [Display(Name = "提示")]
        Prompt = 1,
    }
}
