using Aspose.Words;
using Ninject;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using WPFAIReportCheck.IRepository;


namespace WPFAIReportCheck
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            IKernel kernel = new StandardKernel(new Infrastructure.NinjectDependencyResolver("glz.doc"));

            var ai = kernel.Get<IAIReportCheck>();

            try
            {
                ai.CheckReport();
                MessageBox.Show("已成功校核");
            }
            catch (Exception)
            {

                throw;
            }
            //string fileName = "glz.doc";

            //var doc = new Document(fileName);
            //var ai = new AsposeAIReportCheck(doc);

            //获取word文档中的第一个表格
            //var table0 = doc.GetChildNodes(NodeType.Table, true)[1] as Table;

            //var table0 = ai.GetOverViewTable();
            //var cell = table0.Rows[4].Cells[1];
            //string cbfbm = cell.GetText();

            //string[] splitArray = cbfbm.Split('\r');
            ////用GetText()的方法来获取cell中的值


            //foreach (var s in splitArray)
            //{
            //    s.Replace("\a", "");
            //    s.Replace("\r", "");
            //    Console.WriteLine(s);
            //    Levenshtein(@"《城市桥梁设计规范》（CJJ 11-2011）", @s);
            //};



            //if(ai.reportError!=null)
            //{
            //    foreach(var e in ai.reportError)
            //    {
            //        Console.WriteLine(e.No);
            //        Console.WriteLine(e.Name);
            //    }

            //}
            //Document doc = new Document();

            //Paragraph para1 = new Paragraph(doc);
            //Run run1 = new Run(doc, "Some ");
            //Run run2 = new Run(doc, "text ");
            //para1.AppendChild(run1);
            //para1.AppendChild(run2);
            //doc.FirstSection.Body.AppendChild(para1);

            //Paragraph para2 = new Paragraph(doc);
            //Run run3 = new Run(doc, "is ");
            //Run run4 = new Run(doc, "added ");
            //para2.AppendChild(run3);
            //para2.AppendChild(run4);
            //doc.FirstSection.Body.AppendChild(para2);

            //Comment comment = new Comment(doc, "Awais Hafeez", "AH", DateTime.Today);
            //comment.Paragraphs.Add(new Paragraph(doc));
            //comment.FirstParagraph.Runs.Add(new Run(doc, "Comment text."));

            //CommentRangeStart commentRangeStart = new CommentRangeStart(doc, comment.Id);
            //CommentRangeEnd commentRangeEnd = new CommentRangeEnd(doc, comment.Id);

            //run1.ParentNode.InsertAfter(commentRangeStart, run1);
            //run3.ParentNode.InsertAfter(commentRangeEnd, run3);
            //commentRangeEnd.ParentNode.InsertAfter(comment, commentRangeEnd);
        }
    }
}
