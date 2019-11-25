using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Models;
using OfficeOpenXml;
using System.IO;

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
                FileInfo file = new FileInfo("规范.xlsx");

                using (ExcelPackage package = new ExcelPackage(file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                    int rowCount = 2;// worksheet.Dimension.Rows;   //worksheet.Dimension.Rows指的是所有列中最大行
                    //首行：表头不导入
                    bool rowCur = true;    //行游标指示器
                                           //rowCur=false表示到达行尾
                                           //计算行数
                    while (rowCur)
                    {
                        try
                        {
                            //跳过表头
                            if (string.IsNullOrEmpty(worksheet.Cells[rowCount + 1, 1].Value.ToString()))
                            {
                                rowCur = false;
                            }
                        }
                        catch (Exception)   //读取异常则终止
                        {
                            rowCur = false;
                        }

                        if (rowCur)
                        {
                            rowCount++;
                        }
                    }

                    //bool validationResult = false;
                    int row = 2;    //excel中行指针
                    //行号不为空，则继续添加
                    //while (!string.IsNullOrEmpty(worksheet.Cells[row, 1].Value.ToString()))

                    string combineString = string.Empty;
                    for (row = 2; row <= rowCount; row++)
                    {
                        combineString += worksheet.Cells[row, 2].Value.ToString() + '\r';

                    }
                    Specifications = combineString.Split('\r');

                }

            }
            catch (Exception ex)
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

            //try
            //{
            //    var file = new System.IO.StreamReader("规范.txt", Encoding.Default);
            //    string line;
            //    string combineString = string.Empty;
            //    while ((line = file.ReadLine()) != null)
            //    {
            //        combineString += line + '\r';
            //    }
            //    Specifications = combineString.Split('\r');
            //    file.Close();
            //}
            //catch /*(Exception)*/
            //{
            //    Specifications = new string[]
            //    {
            //        "《城市桥梁设计规范》（CJJ 11-2011）",
            //        "《混凝土结构现场检测技术标准》（GB/T 50784-2013）",
            //        "《公路桥梁荷载试验规程》（JTG/T J21-01-2015）",
            //        "《城市桥梁检测与评定技术规范》（CJJ/T 233-2015）",
            //        "《城市桥梁养护技术标准》（CJJ 99-2017）",
            //    };
            //}

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
