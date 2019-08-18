using System;
using Aspose.Words;
using Aspose.Words.Tables;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Models;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        /// <summary>
        /// 查找报告中静载试验或动载试验中应变或挠度汇总表格中的计算错误
        /// </summary>
        /// <remarks>
        /// 要求：应变和挠度计算表格必须是单位报告模板格式
        /// 算法：遍历应变或挠度检测结果汇总表，假定总应变、弹性应变及满载理论值计算正确，校核残余应变，校验系数及相对残余应变计算是否正确
        /// 测点遍历算法：表格Rows[0].Cells[0]为"测点号"，并且表格Rows[1].Cells[1]为"总应变"或"总变形"
        /// 表格遍历算法：从第3行到最后1行为Cells[1]=总应变/变形，Cells[2]=弹性应变/变形，Cells[4]=满载理论值
        /// 算法：
        /// 已知：总应变，弹性应变
        /// 弹性应变=总应变-残余应变
        /// 校验系数：弹性应变/理论应变
        /// 相对残余应变：残余应变/总应变
        /// TODO：读不出数据时的异常处理
        /// </remarks>
        public void FindStrainOrDispError()
        {
            NodeCollection allTables = _doc.GetChildNodes(NodeType.Table, true);
            for (int i = 0; i < allTables.Count; i++)
            {
                Table table0 = _doc.GetChildNodes(NodeType.Table, true)[i] as Table;
                if (table0.Rows[0].Cells[0].GetText().IndexOf("测点号") >= 0 && table0.Rows[1].Cells[1].GetText().IndexOf("总应变") >= 0)
                {
                    for (int j = 2; j < table0.IndexOf(table0.LastRow); j++)   //TODO：增加最后行尾的判断
                    {
                        try
                        {
                            var totalStrain = Convert.ToDecimal(table0.Rows[j].Cells[1].GetText().Trim().Replace("\a", "").Replace("\r", ""));
                            var elasticStrain = Convert.ToDecimal(table0.Rows[j].Cells[2].GetText().Trim().Replace("\a", "").Replace("\r", ""));
                            var remainStrain = Convert.ToDecimal(table0.Rows[j].Cells[3].GetText().Trim().Replace("\a", "").Replace("\r", ""));
                            var theoryStrain = Convert.ToDecimal(table0.Rows[j].Cells[4].GetText().Trim().Replace("\a", "").Replace("\r", ""));
                            var checkoutCoff = Convert.ToDecimal(table0.Rows[j].Cells[5].GetText().Trim().Replace("\a", "").Replace("\r", ""));
                            decimal relRemainStrain;
                            if (table0.Rows[j].Cells[6].GetText().IndexOf("%") >= 0)
                            {
                                relRemainStrain = Convert.ToDecimal(table0.Rows[j].Cells[6].GetText().Trim().Replace("\a", "").Replace("\r", "").Replace("%", "")) / 100;
                            }
                            else
                            {
                                relRemainStrain = Convert.ToDecimal(table0.Rows[j].Cells[6].GetText().Trim().Replace("\a", "").Replace("\r", "")) / 100;
                            }
                            
                            var calcElasticStrain = totalStrain - remainStrain;
                            var calcCheckoutCoff = Math.Round(calcElasticStrain / theoryStrain, 2);
                            var calcRelRemainStrain = Math.Round(remainStrain / totalStrain, 2);
                            if (calcElasticStrain != elasticStrain)
                            {
                                reportError.Add(new ReportError(ErrorNumber.Calc, $"第{i + 1}张表格", $"计算错误，应为{calcElasticStrain}", true));
                                Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Today);
                                comment.Paragraphs.Add(new Paragraph(_doc));
                                comment.FirstParagraph.Runs.Add(new Run(_doc, $"计算错误，应为{calcElasticStrain}"));
                                DocumentBuilder builder = new DocumentBuilder(_doc);
                                builder.MoveTo(table0.Rows[j].Cells[2].FirstParagraph);
                                builder.CurrentParagraph.AppendChild(comment);
                            }
                            if (calcCheckoutCoff != checkoutCoff)
                            {
                                reportError.Add(new ReportError(ErrorNumber.Calc, $"第{i + 1}张表格", $"计算错误，应为{calcCheckoutCoff}", true));
                                Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Today);
                                comment.Paragraphs.Add(new Paragraph(_doc));
                                comment.FirstParagraph.Runs.Add(new Run(_doc, $"计算错误，应为{calcCheckoutCoff}"));
                                DocumentBuilder builder = new DocumentBuilder(_doc);
                                builder.MoveTo(table0.Rows[j].Cells[5].FirstParagraph);
                                builder.CurrentParagraph.AppendChild(comment);
                            }
                            if (calcRelRemainStrain != relRemainStrain)
                            {
                                reportError.Add(new ReportError(ErrorNumber.Calc, $"第{i + 1}张表格", $"计算错误，应为{$"{calcRelRemainStrain:P}"}", true));
                                Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Today);
                                comment.Paragraphs.Add(new Paragraph(_doc));
                                comment.FirstParagraph.Runs.Add(new Run(_doc, $"计算错误，应为{$"{calcRelRemainStrain:P}"}"));
                                DocumentBuilder builder = new DocumentBuilder(_doc);
                                builder.MoveTo(table0.Rows[j].Cells[6].FirstParagraph);
                                builder.CurrentParagraph.AppendChild(comment);
                            }
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            throw ex;
#else
                            _log.Error($"FindStrainOrDispError函数第{i+1}张表格第{j+1}行数据读取出错，错误信息：{ ex.Message.ToString()}");
                            continue;    //TODO：记录错误
#endif
                        }

                    }
                }
            }
        }
    }
}
