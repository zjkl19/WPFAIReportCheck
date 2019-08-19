using Aspose.Words;
using log4net;
using log4net.Repository;
using Ninject;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.IO;
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
using System.Xml.Linq;
using WPFAIReportCheck.IRepository;
using WPFAIReportCheck.Models;

namespace WPFAIReportCheck
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public ILoggerRepository repository;
        public ILog log;
        public MainWindow()
        {
            InitializeComponent();
            //TODO:movetofunction
            double ScreenWidth = SystemParameters.PrimaryScreenWidth;//WPF
            double ScreenHeight = SystemParameters.PrimaryScreenHeight;//WPF
            Top = 0.3 * (ScreenHeight - Height);
            Left = 0.4 * (ScreenWidth - Width);
            //this.Left = ScreenWidth - this.ActualWidth * 1.3;

            //日志
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);    //防止日志相关问题出现乱码
            repository = LogManager.CreateRepository("WPFAIReportCheck");
            // 默认简单配置，输出至控制台
            //BasicConfigurator.Configure(repository);
            log = LogManager.GetLogger(repository.Name, "WPFAIReportCheckLog4net");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var config = XDocument.Load(@"AIReportCheck.config");
            string doc = string.Empty;
            if (!string.IsNullOrWhiteSpace(FileTextBox.Text))
            {
                doc = FileTextBox.Text;
                IKernel kernel = new StandardKernel(new Infrastructure.NinjectDependencyResolver(doc, log, config));
                var ai = kernel.Get<IAIReportCheck>();

                //注：以上代码相当于以下4行
                //repository = LogManager.CreateRepository("WPFAIReportCheck");
                //log = LogManager.GetLogger(repository.Name, "WPFAIReportCheckLog4net");
                //var doc = @"xxx.doc";
                //var ai = new AsposeAIReportCheck(doc, log);


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
        /// <summary>
        /// 检查更新
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            //ExeConfigurationFileMap map = new ExeConfigurationFileMap
            //{
            //    ExeConfigFilename = @"App1.config"
            //};
            //try
            //{
            //    Configuration config = ConfigurationManager.OpenMappedExeConfiguration(map, ConfigurationUserLevel.None);
            //    //string connstr = config.ConnectionStrings.ConnectionStrings["test1"].ConnectionString;
            //    //MessageBox.Show(connstr);
            //    string key = config.AppSettings.Settings["user"].Value.ToString();
            //    MessageBox.Show(key);
            //}
            //catch (Exception)
            //{
            //    MessageBox.Show("该功能测试中。");
            //    throw;
            //}
            //MessageBox.Show("该功能开发中。");
            //省略了判定xml文件存在
            //XDocument xd = new XDocument();
            var xd = XDocument.Load(@"AIReportCheck.config");
            var xEle1 = xd.Element("configuration").Element("FindDescriptionError").Element("StrainCharactorString");
            MessageBox.Show(xEle1.Name.ToString());
            MessageBox.Show(xEle1.Attribute("version").Value);



        }
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            var doc = "校核结果.docx";
            if (System.IO.File.Exists(doc))
            {
                System.Diagnostics.Process.Start(doc);
            }
        }
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            var doc = "标出错误或警告的报告.doc";
            if (System.IO.File.Exists(doc))
            {
                System.Diagnostics.Process.Start(doc);
            }
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
