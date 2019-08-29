using System;
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
        public void FindSequenceNumberError()
        {
            NodeCollection allTables = _doc.GetChildNodes(NodeType.Table, true);
            for (int i = 0; i < allTables.Count; i++)
            {
                Table table0 = _doc.GetChildNodes(NodeType.Table, true)[i] as Table;
                if (table0.Rows[0].Cells[0].GetText().IndexOf("序号") >= 0)    //含有序号的表格
                {
                    int sn = 1; int convetsn1 = 0;
                    for (int j = 0; j < table0.IndexOf(table0.LastRow); j++)
                    {
                        var sn1 = table0.Rows[j + 1].Cells[0].GetText().Replace("\a", "").Replace("\r", "");
                        try
                        {
                            try    //如果转换出错就当这个序号没有问题
                            {
                                convetsn1 = Convert.ToInt32(sn1);
                            }
                            catch (Exception)
                            {
                                convetsn1 = sn;
                                //TODO:日志
                                //throw;
                            }
                            //TODO:增加转换失败的测试
                            if (convetsn1 != sn)    //转换可能会失败（如序号中含有中文）
                            {
                                reportError.Add(new ReportError(ErrorNumber.Description, $"第{i + 1}张表格", "序号应连贯，从小到大", true));

                                Comment comment = new Comment(_doc, "AI", "AI校核", DateTime.Today);
                                comment.Paragraphs.Add(new Paragraph(_doc));
                                comment.FirstParagraph.Runs.Add(new Run(_doc, "序号应连贯，从小到大"));
                                DocumentBuilder builder = new DocumentBuilder(_doc);
                                builder.MoveTo(table0.Rows[j + 1].Cells[0].FirstParagraph);
                                builder.CurrentParagraph.AppendChild(comment);
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            throw ex;

#else
                            //TODO：增加错误定位信息
                            _log.Error(ex, $"FindSequenceNumberError函数运行出错，错误信息：{ ex.Message.ToString()}");
#endif
                        }

                        sn++;
                    }
                }
            }
        }
    }
}
