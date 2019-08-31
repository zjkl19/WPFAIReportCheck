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

        public ILogger log;
        public MainWindow()
        {
            InitializeComponent();
            //TODO:movetofunction
            Top = 0.3 * (App.ScreenHeight - Height);
            Left = 0.4 * (App.ScreenWidth - Width);
            //this.Left = ScreenWidth - this.ActualWidth * 1.3;

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

        private void StartCheckingButton_Click(object sender, RoutedEventArgs e)
        {
            var config = XDocument.Load(@"AIReportCheck.config");
            string doc = string.Empty;

            if (!string.IsNullOrWhiteSpace(FileTextBox.Text))
            {
                doc = FileTextBox.Text;
                IKernel kernel = new StandardKernel(new Infrastructure.NinjectDependencyResolver(doc, log, config));
                var ai = kernel.Get<IAIReportCheck>();

                #region log ioc
                //注：以上代码相当于以下几行代码
                //LogManager.Configuration = logConfig;
                //log = LogManager.GetCurrentClassLogger();
                //var doc = @"xxx.doc";
                //var config = XDocument.Load(@"AIReportCheck.config");
                //var ai = new AsposeAIReportCheck(doc, log, config);
                #endregion

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
        private void CheckForUpdateButton_Click(object sender, RoutedEventArgs e)
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

            MessageBox.Show("该功能正在开发中");

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

        private void ChooseReportButton_Click(object sender, RoutedEventArgs e)
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
            Environment.Exit(0);
        }

        private void MenuItem_ViewSourceCode_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/zjkl19/WPFAIReportCheck/");
        }

        private void MenuItem_Option_Click(object sender, RoutedEventArgs e)
        {
            var w = new OptionWindow();
            w.Top = 0.4 * (App.ScreenHeight - w.Height);
            w.Left = 0.5 * (App.ScreenWidth - w.Width);
            w.Show();
        }
    }
}
