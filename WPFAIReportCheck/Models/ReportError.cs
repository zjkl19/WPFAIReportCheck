using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using System.Text;

namespace WPFAIReportCheck.Models
{
    class ReportError
    {
        public ErrorNumber No;
        public string Name { get; }
        public string Position { get; }
        public string Description { get; }
        public bool HasComment { get; }
        public ReportError(ErrorNumber No,string Position, string Description,bool HasComment=false)
        {
            this.No = No;
            if (No == ErrorNumber.CMA)
            {
                Name = "计量单位错误";
            }
            else if(No==ErrorNumber.Description)
            {
                Name = "描述错误";
            }
            else
            {
                Name = "待定";
            }

            
            this.Position = Position;
            this.Description = Description;
            this.HasComment = HasComment;
        }
        //def __init__(self, No, Name, Description, Position):      
        //self.No=No
        //self.Name=Name
        //self.Description= Description
        //self.Position= Position
    }
    public enum ErrorNumber
    {
        [Display(Name = "计量错误")]
        CMA = 1,
        [Display(Name = "描述错误")]
        Description = 2,
    }

}
