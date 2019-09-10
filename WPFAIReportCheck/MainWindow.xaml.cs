using Aspose.Words;
using Microsoft.WindowsAPICodePack.Dialogs;
using Ninject;
using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Configuration;
using System.Diagnostics;
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
        BackgroundWorker worker = new BackgroundWorker();
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

            worker.WorkerReportsProgress = true;
            worker.DoWork += new DoWorkEventHandler(BackGroundCheckForUpdate);
            worker.ProgressChanged += BackGroundCheckForUpdate_ProgressChanged;
            worker.RunWorkerAsync();

            //自动更新
            //TODO：配置文件不存在或读取出错
            //            var appConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            //            var autoApdate = Convert.ToBoolean(appConfig.AppSettings.Settings["AutoCheckForUpdate"].Value);

            //            if (autoApdate)
            //            {
            //                AutoCheckForUpdateCheckBox.IsChecked = true;
            //                StatusBarText.Text = "正在检查更新……";
            //            }
            //            else
            //            {
            //                AutoCheckForUpdateCheckBox.IsChecked = false;
            //            }

            //            new Thread(() =>
            //            {
            //                Dispatcher.BeginInvoke(new Action(() =>
            //                {

            //                    try
            //                    {
            //                        if (autoApdate)
            //                        {
            //                            Repository.CheckForUpdate.CheckByRestClient();
            //                            StatusBarText.Text ="就绪";
            //                        }

            //                    }
            //                    catch (Exception ex)
            //                    {
            //#if DEBUG
            //                        throw ex;

            //#else
            //                        log.Error(ex, $"\"自动校核\"运行出错，错误信息：{ ex.Message.ToString()}");
            //#endif
            //                    }
            //                }));
            //            }).Start();

        }
        /// <summary>
        /// 进度返回处理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void BackGroundCheckForUpdate_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //var appConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            //var autoApdate = Convert.ToBoolean(appConfig.AppSettings.Settings["AutoCheckForUpdate"].Value);

            StatusBarText.Text = e.UserState.ToString();

            //if (autoApdate)
            //{
            //    AutoCheckForUpdateCheckBox.IsChecked = true;
            //    StatusBarText.Text = "正在检查更新……";
            //}
            //else
            //{
            //    AutoCheckForUpdateCheckBox.IsChecked = false;
            //}
        }

        void BackGroundCheckForUpdate(object sender, DoWorkEventArgs e)
        {
            //for (int i = 0; i <= 100; i++)
            //{
            //    worker.ReportProgress(i);//返回进度
            //    Thread.Sleep(100);
            //}
            //自动更新
            //TODO：配置文件不存在或读取出错

            
            var appConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            var autoApdate = Convert.ToBoolean(appConfig.AppSettings.Settings["AutoCheckForUpdate"].Value);
            
            worker.ReportProgress(1, "正在检查更新……");

            //if (autoApdate)
            //{
            //    AutoCheckForUpdateCheckBox.IsChecked = true;
            //    StatusBarText.Text = "正在检查更新……";
            //}
            //else
            //{
            //    AutoCheckForUpdateCheckBox.IsChecked = false;
            //}

            try
            {
                if (autoApdate)
                {
                    Repository.CheckForUpdate.CheckByRestClient();
                    
                    //StatusBarText.Text = "就绪";
                    worker.ReportProgress(1, "就绪");
                    
                }

            }
            catch (Exception ex)
            {
#if DEBUG
                throw ex;

#else
                log.Error(ex, $"\"自动校核\"运行出错，错误信息：{ ex.Message.ToString()}");
#endif
            }
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

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            var doc = "校核结果.docx";
            if (File.Exists(doc))
            {
                Process.Start(doc);
            }
            else
            {
                MessageBox.Show("请先校核");
            }
        }
        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            var doc = "标出错误或警告的报告.doc";
            if (File.Exists(doc))
            {
                Process.Start(doc);
            }
            else
            {
                MessageBox.Show("请先校核");
            }
        }

        private void ChooseReportButton_Click(object sender, RoutedEventArgs e)
        {
            using (
                var dialog = new CommonOpenFileDialog
                {
                    IsFolderPicker = false,//设置为选择文件夹
                    DefaultDirectory = @"\",

                })
            {
                dialog.Filters.Add(new CommonFileDialogFilter("Word 文档", "*.docx;*.doc"));

                //如果要分开过滤，考虑以下代码
                //dialog.Filters.Add(new CommonFileDialogFilter("Word 97-2003 文档", "*.doc"));

                var result = dialog.ShowDialog();
                if (result == CommonFileDialogResult.Ok)
                {
                    FileTextBox.Text = dialog.FileName;
                };
            }

        }

        private void InstructionsButton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("该功能开发中");
        }

        private void AutoCheckForUpdateCheckBox_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                var appConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                if (AutoCheckForUpdateCheckBox.IsChecked ?? false)
                {
                    appConfig.AppSettings.Settings["AutoCheckForUpdate"].Value = "true";
                }
                else
                {
                    appConfig.AppSettings.Settings["AutoCheckForUpdate"].Value = "false";
                }

                appConfig.Save(ConfigurationSaveMode.Modified);

                ConfigurationManager.RefreshSection("appSettings");
            }
            catch (Exception ex)
            {
#if DEBUG
                throw ex;

#else
                log.Error(ex, $"\"自动校核\"运行出错，错误信息：{ ex.Message.ToString()}");
#endif
            }

        }
    }
}
