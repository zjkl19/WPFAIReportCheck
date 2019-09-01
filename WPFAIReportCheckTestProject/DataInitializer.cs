using System;
using System.Collections.Generic;
using System.Text;

namespace WPFAIReportCheckTestProject
{
    public class DataInitializer
    {
        public static string AIReportCheckXml = @"<?xml version=""1.0"" encoding=""utf - 8"" ?>
                            <configuration>
                              <FindStrainOrDispError row1=""0"" col1=""0"" row2 =""1"" col2 =""1""  charactorString =""测点号"" >
                                <Strain charactorString = ""总应变"" />
                                <Disp charactorString = ""总变形"" />
                             </FindStrainOrDispError >
                           </configuration >";
    }
}
