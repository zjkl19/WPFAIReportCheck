using Aspose.Words;
using Ninject;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
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
using System.Windows.Threading;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Models;

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

            string doc = string.Empty;
            if (!string.IsNullOrWhiteSpace(FileTextBox.Text))
            {
                doc = FileTextBox.Text;

                IKernel kernel = new StandardKernel(new Infrastructure.NinjectDependencyResolver(doc));

                var ai = kernel.Get<IAIReportCheck>();
                //Dispatcher.BeginInvoke(DispatcherPriority.SystemIdle, new Action(() =>
                //{
                //    try
                //    {
                //        ai.CheckReport();
                //        MessageBox.Show("已成功校核");
                //    }
                //    catch (Exception)
                //    {
                //        throw;
                //    }
                //}));
                new Thread(() =>
                {
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            ai.CheckReport();
                            MessageBox.Show("已成功校核");
                        }
                        catch (Exception)
                        {
                            throw;
                        }
                    }));
                }).Start();
            }
            else
            {
                MessageBox.Show("请输入文件名");
            }
            //try
            //{
            //    ai.CheckReport();
            //    MessageBox.Show("已成功校核");
            //}
            //catch (Exception)
            //{

            //    throw;
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

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("校核结果仅供参考");
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            MessageBox.Show(GetEnumDesc(ErrorNumber.CMA).ToString());
        }
        public static string GetEnumDesc(Enum en)
        {
            Type type = en.GetType();
            var memInfo = type.GetMember(en.ToString());
            if (memInfo != null && memInfo.Length > 0)
            {
                object[] attrs = memInfo[0].GetCustomAttributes(typeof(System.ComponentModel.DataAnnotations.DisplayAttribute), false); if (attrs != null && attrs.Length > 0) return ((System.ComponentModel.DataAnnotations.DisplayAttribute)attrs[0]).Name;
            }
            return en.ToString();
        }
    }
    
    //public static class EnumHelper<T>
    //{
    //    public static string GetNameFromValue(T value)
    //    {
    //        var type = typeof(T);
    //        if (!type.IsEnum) throw new InvalidOperationException();

    //        foreach (var field in type.GetFields())
    //        {
    //            var attribute = Attribute.GetCustomAttribute(field,
    //                typeof(DisplayAttribute)) as DisplayAttribute;
    //            if (((T)field.GetValue(null)).ToString()==value.ToString())
    //            {

    //                return attribute.Name;

    //            }
    //            else
    //            {
    //                continue;
    //            }
    //        }

    //        throw new ArgumentOutOfRangeException("name");
    //    }
    //}
    //public static class EnumHelper1<T>
    //{
    //    public static T GetValueFromName(string name)
    //    {
    //        var type = typeof(T);
    //        if (!type.IsEnum) throw new InvalidOperationException();

    //        foreach (var field in type.GetFields())
    //        {
    //            var attribute = Attribute.GetCustomAttribute(field,
    //                typeof(DisplayAttribute)) as DisplayAttribute;
    //            if (attribute != null)
    //            {
    //                if (attribute.Name == name)
    //                {
    //                    return (T)field.GetValue(null);
    //                }
    //            }
    //            else
    //            {
    //                if (field.Name == name)
    //                    return (T)field.GetValue(null);
    //            }
    //        }

    //        throw new ArgumentOutOfRangeException("name");
    //    }
    //}
}
