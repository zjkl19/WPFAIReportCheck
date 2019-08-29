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
            var ScreenWidth = SystemParameters.PrimaryScreenWidth;//WPF
            var ScreenHeight = SystemParameters.PrimaryScreenHeight;//WPF
            var config = XDocument.Load(@"Option.config");
            var op = Convert.ToInt32(config.Element("configuration").Element("FindUnitError").Value);
            var w = new ProgressBarWindow();
            w.Top = 0.4 * (ScreenHeight - w.Height);
            w.Left = 0.4 * (ScreenWidth - w.Width);
            w.Show();
            var thread = new Thread(new ThreadStart(() =>
            {
                if (op == 1)
                {
                    _FindUnitError();
                }
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 10; });
                Thread.Sleep(100);
                _FindNotExplainComponentNo();
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 20; });
                Thread.Sleep(100);
                _FindSpecificationsError();

                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 30; });
                Thread.Sleep(100);
                _FindSequenceNumberError();
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 40; });
                Thread.Sleep(100);
                FindStrainOrDispError();
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 50; });
                Thread.Sleep(100);
                FindDescriptionError();
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 60; });
                Thread.Sleep(100);
                FindOtherBridgesWarnning();
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 70; });
                Thread.Sleep(100);

                _GenerateResultReport();
                _doc.Save("标出错误或警告的报告.doc");
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 100; });
                Thread.Sleep(100);
                MessageBox.Show("已成功校核");

            }));
            thread.Start();

            w.Close();



        }
    }
}
