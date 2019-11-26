using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using OfficeOpenXml;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Models;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        /// <summary>
        /// 查找报告中名词描述错误
        /// </summary>
        /// <remarks>
        /// 算法：用正则表达式全文搜索“CNAS检查机构”关键字，如果查到，表示有误
        /// TODO：增加其它名词描述错误
        /// </remarks>
        /// 算法主要参与人员：林迪南
        public void FindDescriptionError()
        {
            FindReplaceOptions options;
            MatchCollection matches;

            var regexList = new List<Regex>();    //错误的描述
            var correctList = new List<string>();    //正确的描述

            try
            {
                var file = new FileInfo("描述错误.xlsx");

                using (var package = new ExcelPackage(file))
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

                    for (row = 2; row <= rowCount; row++)
                    {
                        regexList.Add(new Regex($"{ worksheet.Cells[row, 2].Value.ToString() }"));
                        correctList.Add(worksheet.Cells[row, 3].Value.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                regexList.Clear();
                regexList.Add(new Regex(@"CNAS检查机构"));
                correctList.Clear();
                correctList.Add("CNAS检验机构");
            }

            Regex regex; 

            for (int i=0;i<regexList.Count;i++)
            {
                regex = regexList[i];
                try
                {
                    matches = regex.Matches(_originalWholeText);

                    if (matches.Count != 0)
                    {
                        foreach (Match m in matches)
                        {
                            reportError.Add(new ReportError(ErrorNumber.Description, "正文" + m.Index.ToString(), $"应为\"{ correctList[i] }\"", true));
                        }

                        options = new FindReplaceOptions
                        {
                            ReplacingCallback = new ReplaceEvaluatorFindAndHighlightWithComment(_doc, "AI校核", $"应为\"{ correctList[i] }\""),
                            Direction = FindReplaceDirection.Forward
                        };
                        _doc.Range.Replace(regex, "", options);

                    }
                }
                catch (Exception ex)
                {
#if DEBUG
                throw ex;

#else
                    _log.Error(ex, $"FindDescriptionError函数运行出错，错误信息：{ ex.Message.ToString()}");
#endif
                }
            }

            
//            try
//            {
//                matches = regex.Matches(_originalWholeText);

//                if (matches.Count != 0)
//                {
//                    foreach (Match m in matches)
//                    {
//                        reportError.Add(new ReportError(ErrorNumber.Description, "正文" + m.Index.ToString(), "应为\"CNAS检验机构\"", true));
//                    }

//                    options = new FindReplaceOptions
//                    {
//                        ReplacingCallback = new ReplaceEvaluatorFindAndHighlightWithComment(_doc, "AI校核", "应为\"CNAS检验机构\""),
//                        Direction = FindReplaceDirection.Forward
//                    };
//                    _doc.Range.Replace(regex, "", options);

//                }
//            }
//            catch (Exception ex)
//            {
//#if DEBUG
//                throw ex;

//#else
//                            _log.Error(ex, $"FindDescriptionError函数运行出错，错误信息：{ ex.Message.ToString()}");
//#endif
//            }
 
        }
    }
}
