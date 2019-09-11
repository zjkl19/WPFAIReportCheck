using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Tables;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Models;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        /// <summary>
        /// 查找报告中静载试验中应变或挠度汇总表格中的计算错误
        /// </summary>
        /// <remarks>
        /// 要求：应变和挠度计算表格必须是单位报告模板格式
        /// 算法：遍历应变或挠度检测结果汇总表，假定总变形、弹性变形及满载理论值计算正确，校核残余变形，校验系数及相对残余变形计算是否正确
        /// 测点遍历算法：表格Rows[0].Cells[0]为"测点号"，并且表格Rows[1].Cells[1]为"总应变"或"总变形"
        /// 表格遍历算法：从第3行到最后1行为Cells[1]=总变形，Cells[2]=弹性应变/变形，Cells[4]=满载理论值
        /// 算法：
        /// 已知：总变形，弹性变形
        /// 弹性变形=总变形-残余变形
        /// 校验系数：弹性变形/理论变形
        /// 相对残余变形：残余变形/总变形
        /// TODO：读不出数据时的异常处理，要能定位出具体位置
        /// </remarks>
        /// 算法主要参与人员：林迪南、陈思远
        public void FindStrainOrDispError()
        {

            int row1,col1,row2,col2;
            string headerCharactorString,strainCharactorString,dispCharactorString;
            Regex regexHeader;
            try
            {
                //var xEle1 = xd.Element("configuration").Element("FindDescriptionError").Element("StrainCharactorString");
                //MessageBox.Show(xEle1.Name.ToString());
                //MessageBox.Show(xEle1.Attribute("version").Value);
                headerCharactorString=_config.Element("configuration").Element("FindStrainOrDispError").Attribute("charactorString").Value;
                row1 =Convert.ToInt32(_config.Element("configuration").Element("FindStrainOrDispError").Attribute("row1").Value);
                col1 = Convert.ToInt32(_config.Element("configuration").Element("FindStrainOrDispError").Attribute("col1").Value);
                row2 = Convert.ToInt32(_config.Element("configuration").Element("FindStrainOrDispError").Attribute("row2").Value);
                col2 = Convert.ToInt32(_config.Element("configuration").Element("FindStrainOrDispError").Attribute("col2").Value);
                strainCharactorString = _config.Element("configuration").Element("FindStrainOrDispError").Element("Strain").Attribute("charactorString").Value;
                dispCharactorString = _config.Element("configuration").Element("FindStrainOrDispError").Element("Disp").Attribute("charactorString").Value;
            }
            catch (Exception)
            {
                row1 = 0;col1 = 0;row2 = 1;col2 = 1;
                headerCharactorString = @"测点[号]?"; strainCharactorString = "总应变"; dispCharactorString = "总变形";
                //TODO：增加条件编译的异常处理
            }

            regexHeader = new Regex(headerCharactorString);

            int tableLastRow = 0;
            NodeCollection allTables = _originalDoc.GetChildNodes(NodeType.Table, true);
            for (int i = 0; i < allTables.Count; i++)
            {
                Table table0 = _originalDoc.GetChildNodes(NodeType.Table, true)[i] as Table;

                Table table1 = _doc.GetChildNodes(NodeType.Table, true)[i] as Table;    //要写入批注的文档

                if (regexHeader.Matches(table0.Rows[row1].Cells[col1].GetText()).Count > 0
                     && (table0.Rows[row2].Cells[col2].GetText().IndexOf(strainCharactorString) >= 0 || table0.Rows[row2].Cells[col2].GetText().IndexOf(dispCharactorString) >= 0))
                {
                    tableLastRow = table0.IndexOf(table0.LastRow);
                    if (table0.Rows[table0.IndexOf(table0.LastRow)].Cells.Count <= 2)    //最后1行单元格个数不超过2个
                    {
                        tableLastRow = table0.IndexOf(table0.LastRow) - 1;
                    }
                    for (int j = 2; j < tableLastRow; j++)   //TODO：增加最后行尾的判断
                    {
                        try
                        {
                            var totalDeform = Convert.ToDecimal(table0.Rows[j].Cells[1].GetText().Trim().Replace("\a", "").Replace("\r", ""));
                            var elasticDeform = Convert.ToDecimal(table0.Rows[j].Cells[2].GetText().Trim().Replace("\a", "").Replace("\r", ""));
                            var remainDeform = Convert.ToDecimal(table0.Rows[j].Cells[3].GetText().Trim().Replace("\a", "").Replace("\r", ""));
                            var theoryDeform = Convert.ToDecimal(table0.Rows[j].Cells[4].GetText().Trim().Replace("\a", "").Replace("\r", ""));
                            var checkoutCoff = Convert.ToDecimal(table0.Rows[j].Cells[5].GetText().Trim().Replace("\a", "").Replace("\r", ""));
                            decimal relRemainDeform;
                            if (table0.Rows[j].Cells[6].GetText().IndexOf("%") >= 0)
                            {
                                relRemainDeform = Convert.ToDecimal(table0.Rows[j].Cells[6].GetText().Trim().Replace("\a", "").Replace("\r", "").Replace("%", "")) / 100;
                            }
                            else
                            {
                                relRemainDeform = Convert.ToDecimal(table0.Rows[j].Cells[6].GetText().Trim().Replace("\a", "").Replace("\r", "")) / 100;
                            }
                            
                            var calcElasticDeform = totalDeform - remainDeform;
                            var calcCheckoutCoff = Math.Round(calcElasticDeform / theoryDeform, 2);
                            var calcRelRemainDeform = Math.Round(remainDeform / totalDeform, 4);
                            if (calcElasticDeform != elasticDeform)
                            {
                                reportError.Add(new ReportError(ErrorNumber.Calc, $"第{i + 1}张表格", $"计算错误，应为{calcElasticDeform}", true));
                                Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Today);
                                comment.Paragraphs.Add(new Paragraph(_doc));
                                comment.FirstParagraph.Runs.Add(new Run(_doc, $"计算错误，应为{calcElasticDeform}"));
                                DocumentBuilder builder = new DocumentBuilder(_doc);
                                builder.MoveTo(table1.Rows[j].Cells[2].FirstParagraph);
                                builder.CurrentParagraph.AppendChild(comment);
                            }
                            if (calcCheckoutCoff != checkoutCoff)
                            {
                                reportError.Add(new ReportError(ErrorNumber.Calc, $"第{i + 1}张表格", $"计算错误，应为{calcCheckoutCoff}", true));
                                Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Today);
                                comment.Paragraphs.Add(new Paragraph(_doc));
                                comment.FirstParagraph.Runs.Add(new Run(_doc, $"计算错误，应为{calcCheckoutCoff}"));
                                DocumentBuilder builder = new DocumentBuilder(_doc);
                                builder.MoveTo(table1.Rows[j].Cells[5].FirstParagraph);
                                builder.CurrentParagraph.AppendChild(comment);
                            }
                            if (calcRelRemainDeform != relRemainDeform)
                            {
                                reportError.Add(new ReportError(ErrorNumber.Calc, $"第{i + 1}张表格", $"计算错误，应为{$"{calcRelRemainDeform:P}"}", true));
                                Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Today);
                                comment.Paragraphs.Add(new Paragraph(_doc));
                                comment.FirstParagraph.Runs.Add(new Run(_doc, $"计算错误，应为{$"{calcRelRemainDeform:P}"}"));
                                DocumentBuilder builder = new DocumentBuilder(_doc);
                                builder.MoveTo(table1.Rows[j].Cells[6].FirstParagraph);
                                builder.CurrentParagraph.AppendChild(comment);
                            }
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            throw ex;
                            
#else
                            _log.Error(ex,$"FindDeformOrDispError函数第{i+1}张表格第{j+1}行数据读取出错，错误信息：{ ex.Message.ToString()}");
                            continue;    //TODO：记录错误
#endif
                        }

                    }
                }
            }
        }
    }
}
