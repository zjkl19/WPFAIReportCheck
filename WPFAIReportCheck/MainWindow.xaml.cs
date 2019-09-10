using Microsoft.WindowsAPICodePack.Dialogs;
using NLog;
using System;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Windows;


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

            Title = $"报告自动校核 v{Application.ResourceAssembly.GetName().Version.ToString()}";

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
            CheckForUpdateInStarup();    //启动时检查更新
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
