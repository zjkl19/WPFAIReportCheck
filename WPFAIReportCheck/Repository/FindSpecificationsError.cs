using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Models;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        //20190825Bug:列表最后一个规范可能校核不出来
        public void FindSpecificationsError()
        {
            string[] Specifications;
            //规范
            try
            {
                var file = new System.IO.StreamReader("规范.txt", Encoding.Default);
                string line;
                string combineString = string.Empty;
                while ((line = file.ReadLine()) != null)
                {
                    combineString += line + '\r';
                }
                Specifications = combineString.Split('\r');
                file.Close();
            }
            catch (Exception)
            {
                Specifications = new string[]
                {
                    "《城市桥梁设计规范》（CJJ 11-2011）",
                    "《混凝土结构现场检测技术标准》（GB/T 50784-2013）",
                    "《公路桥梁荷载试验规程》（JTG/T J21-01-2015）",
                    "《城市桥梁检测与评定技术规范》（CJJ/T 233-2015）",
                    "《城市桥梁养护技术标准》（CJJ 99-2017）",
                };
            }

            double similarity;    //相似度
                                  //获取word文档中的第一个表格

            //Table table0=null;
            //Cell cell=null;
            //FindReplaceOptions options;
            //string repStr=string.Empty;
            double similarityLBound = 0.85;
            double similarityUBound = 1.00;

            NodeCollection allTables = _doc.GetChildNodes(NodeType.Table, true);
            for (int i = 0; i < allTables.Count; i++)
            {
                Table table0 = _doc.GetChildNodes(NodeType.Table, true)[i] as Table;
                if ((table0.Rows[0].Cells[0].GetText().IndexOf("委托单位") >= 0))
                {
                    Cell cell = table0.Rows[4].Cells[1];
                    string[] splitArray = cell.GetText().Split('\r');    //用GetText()的方法来获取cell中的值
                    foreach (var s in splitArray)
                    {
                        string repStr = s;
                        repStr = repStr.Replace("\a", "").Replace("\r", "").Replace("(", "（").Replace(")", "）");
                        var s1 = Regex.Replace(repStr, @"(.+)《", "《");    //替换"《"之前的内容为""
                        foreach (var sp in Specifications)
                        {
                            similarity = Levenshtein(@s1, @sp);
                            if (similarity > similarityLBound && similarity < similarityUBound)
                            {
                                var regex = new Regex(@s.Replace("\a", ""));
                                reportError.Add(new ReportError(ErrorNumber.Description, "汇总表格中主要检测检验依据", "应为" + sp, true));
                                FindReplaceOptions options = new FindReplaceOptions
                                {
                                    ReplacingCallback = new ReplaceEvaluatorFindAndHighlightWithComment(_doc, "AI校核", "应为" + sp),
                                    Direction = FindReplaceDirection.Forward
                                };
                                _doc.Range.Replace(regex.ToString(), "", options);  //注意！直接用regex可能会出错，应为()等字符是元字符

                                break;
                            }
                        }
                    };
                    break;
                }
            }
            //var table0 = ai.GetOverViewTable();
        }

    }

}
