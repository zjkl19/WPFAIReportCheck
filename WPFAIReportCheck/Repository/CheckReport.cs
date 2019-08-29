using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;
using WPFAIReportCheck.IRepository;

namespace WPFAIReportCheck.Repository
{
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        public void CheckReport()
        {

            var config = XDocument.Load(@"Option.config");
            var op = Convert.ToInt32(config.Element("configuration").Element("FindUnitError").Value);

            var w = new ProgressBarWindow();
            w.Top = 0.4 * (App.ScreenHeight - w.Height);
            w.Left = 0.4 * (App.ScreenWidth - w.Width);

            var progressSleepTime = 100;

            var thread = new Thread(new ThreadStart(() =>
            {
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.Show(); });
                _FindUnitError();
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 10; });
                Thread.Sleep(progressSleepTime);
                _FindNotExplainComponentNo();
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 20; });
                Thread.Sleep(progressSleepTime);
                _FindSpecificationsError();

                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 30; });
                Thread.Sleep(progressSleepTime);
                FindSequenceNumberError();
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 40; });
                Thread.Sleep(progressSleepTime);
                FindStrainOrDispError();
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 50; });
                Thread.Sleep(progressSleepTime);
                FindDescriptionError();
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 60; });
                Thread.Sleep(progressSleepTime);
                FindOtherBridgesWarnning();
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 70; });
                Thread.Sleep(progressSleepTime);

                _GenerateResultReport();
                _doc.Save("标出错误或警告的报告.doc");
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 100; });
                Thread.Sleep(100);
                MessageBox.Show("已成功校核");
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.Close(); });
            }));
            thread.Start();
            


        }
    }
}
