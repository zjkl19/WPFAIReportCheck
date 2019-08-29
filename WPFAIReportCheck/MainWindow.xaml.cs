using Aspose.Words;
using Microsoft.WindowsAPICodePack.Dialogs;
//using log4net;
//using log4net.Repository;
using Ninject;
using NLog;
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
        //public ILoggerRepository repository;
        //public ILog log;
        public ILogger log;
        public double ScreenWidth = SystemParameters.PrimaryScreenWidth;//WPF
        public double ScreenHeight = SystemParameters.PrimaryScreenHeight;//WPF
        public MainWindow()
        {
            InitializeComponent();
            //TODO:movetofunction
            //ScreenWidth = SystemParameters.PrimaryScreenWidth;
            //ScreenHeight = SystemParameters.PrimaryScreenHeight;
            Top = 0.3 * (ScreenHeight - Height);
            Left = 0.4 * (ScreenWidth - Width);
            //this.Left = ScreenWidth - this.ActualWidth * 1.3;

            //日志
            //log4net
            //Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);    //防止日志相关问题出现乱码
            //repository = LogManager.CreateRepository("WPFAIReportCheck");
            // 默认简单配置，输出至控制台
            //BasicConfigurator.Configure(repository);
            //log = LogManager.GetLogger(repository.Name, "WPFAIReportCheckLog4net");

            //Nlog
            var config = new NLog.Config.LoggingConfiguration();

            // Targets where to log to: File and Console
            var logfile = new NLog.Targets.FileTarget("logfile") { FileName = @"Log\LogFile.txt" };
            var logconsole = new NLog.Targets.ConsoleTarget("logconsole");

            // Rules for mapping loggers to targets            
            config.AddRule(LogLevel.Info, LogLevel.Fatal, logconsole);
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, logfile);

            // Apply config           
            LogManager.Configuration = config;

            log = LogManager.GetCurrentClassLogger();
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
                            
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            throw ex;

#else
                            log.Error(ex, $"\"自动校核\"运行出错，错误信息：{ ex.Message.ToString()}");
#endif
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

        //免责声明
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

            //以下为临时测试代码
            //try
            //{
            //    var xd = XDocument.Load(@"AIReportCheck.config");
            //    var xEle1 = xd.Element("configuration").Element("FindDescriptionError").Element("StrainCharactorString");
            //    MessageBox.Show(xEle1.Name.ToString());
            //    MessageBox.Show(xEle1.Attribute("version").Value);
            //}
            //catch (Exception)
            //{

            //    throw;
            //}

            ProgressBegin();


        }

        private void ProgressBegin()
        {
            //var w = new ProgressBarWindow()
            //{
            //    Top = 0.4 * (ScreenHeight - Height),
            //    Left = 0.5 * (ScreenWidth - Width),
            //};
            //w.Show();
            //Thread thread = new Thread(new ThreadStart(() =>
            //{
            //    for (int i = 0; i <= 100; i++)
            //    {
            //        w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = i; });
            //        Thread.Sleep(100);
            //    }

            //}));
            //thread.Start();
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

        private void OpenFileDialogButton_Click(object sender, RoutedEventArgs e)
        {
            using (
                var dialog = new CommonOpenFileDialog
                {
                    IsFolderPicker = false,//设置为选择文件夹
                    DefaultDirectory= @"\",
                    
                })
            {
                dialog.Filters.Add(new CommonFileDialogFilter("Word 文档", "*.docx;*.doc"));
                //dialog.Filters.Add(new CommonFileDialogFilter("Word 97-2003 文档", "*.doc"));

                var result = dialog.ShowDialog();
                if (result == CommonFileDialogResult.Ok)
                {
                    FileTextBox.Text = dialog.FileName;
                };
            }

        }

        private void MenuItem_About_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("系统框架设计及编程：路桥监测研究所林迪南、陈思远等");
        }

        private void MenuItem_Exit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void MenuItem_ViewSourceCode_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/zjkl19/WPFAIReportCheck/");
        }

        private void MenuItem_Option_Click(object sender, RoutedEventArgs e)
        {
            var w = new OptionWindow
            {
                Top = 0.4 * (ScreenHeight - Height),
                Left = 0.5 * (ScreenWidth - Width),
            };
            w.Show();
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
