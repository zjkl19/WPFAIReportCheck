using System;
using System.Collections.Generic;
using System.Text;

namespace WPFAIReportCheck.Models
{
    class ReportError
    {
        public ErrorNumber No;
        public string Name { get; }

        public string Position { get; }

        public string Description { get; }
        public ReportError(ErrorNumber No,string Position, string Description)
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
        }
        //def __init__(self, No, Name, Description, Position):      
        //self.No=No
        //self.Name=Name
        //self.Description= Description
        //self.Position= Position
    }

    public enum ErrorNumber
    {
        CMA = 1,
        Description=2,
    }
}
