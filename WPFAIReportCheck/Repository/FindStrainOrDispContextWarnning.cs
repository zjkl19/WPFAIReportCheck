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
        /// 要求：应变和挠度计算表格必须是单位报告模板格式，结果统计模板为单位报告的模板
        /// 算法：遍历应变或挠度检测结果汇总表，假定表格中各参数计算正确，程序对各参数进行统计，
        /// 然后查找上文及"检验结果"汇总表中是否含有对应统计信息，如未发现，则发出警告。
        /// 测点遍历算法：表格Rows[0].Cells[0]为"测点号"，并且表格Rows[1].Cells[1]为"总应变"或"总变形"
        /// 表格遍历算法：从第3行到最后1行为Cells[1]=总变形，Cells[2]=弹性应变/变形，Cells[4]=满载理论值
        /// 算法：
        /// 表格上文文字查找算法：table.PreviousSibling.PreviousSibling.Range.Text
        /// TODO：读不出数据时的异常处理，要能定位出具体位置
        /// </remarks>
        /// 算法主要参与人员：林迪南、陈思远
        public void FindStrainOrDispContextWarnning()
        {

            int row1, col1, row2, col2;
            string headerCharactorString, strainCharactorString, dispCharactorString;
            try
            {
                headerCharactorString = _config.Element("configuration").Element("FindStrainOrDispError").Attribute("charactorString").Value;
                row1 = Convert.ToInt32(_config.Element("configuration").Element("FindStrainOrDispError").Attribute("row1").Value);
                col1 = Convert.ToInt32(_config.Element("configuration").Element("FindStrainOrDispError").Attribute("col1").Value);
                row2 = Convert.ToInt32(_config.Element("configuration").Element("FindStrainOrDispError").Attribute("row2").Value);
                col2 = Convert.ToInt32(_config.Element("configuration").Element("FindStrainOrDispError").Attribute("col2").Value);
                strainCharactorString = _config.Element("configuration").Element("FindStrainOrDispError").Element("Strain").Attribute("charactorString").Value;
                dispCharactorString = _config.Element("configuration").Element("FindStrainOrDispError").Element("Disp").Attribute("charactorString").Value;
            }
            catch (Exception)
            {
                row1 = 0; col1 = 0; row2 = 1; col2 = 1;
                headerCharactorString = "测点号"; strainCharactorString = "总应变"; dispCharactorString = "总变形";
                //TODO：增加条件编译的异常处理
            }

            int tableLastRow = 0;

            NodeCollection allTables = _doc.GetChildNodes(NodeType.Table, true);
            for (int i = 0; i < allTables.Count; i++)
            {
                decimal maxElasticDeform = 0.0m;
                decimal minCheckoutCoff = decimal.MaxValue; decimal maxCheckoutCoff = 0.0m;
                decimal minRelRemainDeform = decimal.MaxValue; decimal maxRelRemainDeform = 0.0m;
                Table table0 = _doc.GetChildNodes(NodeType.Table, true)[i] as Table;
                try
                {
                    if (table0.Rows[row1].Cells[col1].GetText().IndexOf(headerCharactorString) >= 0
                     && (table0.Rows[row2].Cells[col2].GetText().IndexOf(strainCharactorString) >= 0 || table0.Rows[row2].Cells[col2].GetText().IndexOf(dispCharactorString) >= 0))
                    {
                        tableLastRow = table0.IndexOf(table0.LastRow);
                        if (table0.Rows[table0.IndexOf(table0.LastRow)].Cells[0].GetText().Contains("备注"))    //如果最后一行含有备注，遍历的行要减1
                        {
                            tableLastRow = table0.IndexOf(table0.LastRow) - 1;
                        }
                        for (int j = 2; j < tableLastRow; j++)   //TODO：增加最后行尾的判断
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
                            maxElasticDeform = elasticDeform > maxElasticDeform ? elasticDeform : maxElasticDeform;
                            minCheckoutCoff = checkoutCoff < minCheckoutCoff ? checkoutCoff : minCheckoutCoff;
                            maxCheckoutCoff = checkoutCoff > maxCheckoutCoff ? checkoutCoff : maxCheckoutCoff;
                            minRelRemainDeform = relRemainDeform < minRelRemainDeform ? relRemainDeform : minRelRemainDeform;
                            maxRelRemainDeform = relRemainDeform > maxRelRemainDeform ? relRemainDeform : maxRelRemainDeform;
                        }

                        var ctx = table0.PreviousSibling.PreviousSibling.Range.Text;
                        MatchCollection matches;
                        var regex = new Regex($"{maxElasticDeform}");
                        matches = regex.Matches(ctx);

                        if (matches.Count == 0)
                        {

                            reportWarnning.Add(new ReportWarnning(WarnningNumber.NotFoundInfo, $"第{i + 1}张表格", $"关键字{maxElasticDeform}未找到", true));

                            Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Now);
                            comment.Paragraphs.Add(new Paragraph(_doc));
                            comment.FirstParagraph.Runs.Add(new Run(_doc, $"关键字{maxElasticDeform}未找到"));
                            DocumentBuilder builder = new DocumentBuilder(_doc);
                            builder.MoveTo(table0.PreviousSibling.PreviousSibling);
                            builder.CurrentParagraph.AppendChild(comment);
                        }
                    }


                }
                catch (Exception ex)
                {
#if DEBUG
                    throw ex;

#else
                    _log.Error(ex, $"FindStrainOrDispContextWarnning函数第{i + 1}张表格出错，错误信息：{ ex.Message.ToString()}");
                    continue;    //TODO：记录错误
#endif
                }

              

            }
        }
    }
}

