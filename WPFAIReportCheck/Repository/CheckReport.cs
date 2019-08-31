using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Xml.Linq;
using WPFAIReportCheck.IRepository;

namespace WPFAIReportCheck.Repository
{
    public class ProgressBarDataBinding : INotifyPropertyChanged
    {

        private int v = 0;
        public int V { get
            { return v; } set
            {
                v = value;

                OnPropertyChanged("V");
            }
        }
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
    public partial class AsposeAIReportCheck : IAIReportCheck
    {
        private delegate void MethodDelegate();
        public void CheckReport()
        {

            //举例：
            //SelectListInt =[1,0]
            //SelectListFunctionName=[FindUnitError,FindNotExplainComponentNo]
            //表示算法只算FindUnitError,不算FindNotExplainComponentNo
            var controlClass = new
            {
                SelectListInt = new List<int>(),
                SelectListFunctionName = new List<MethodDelegate> {
                    _FindUnitError,_FindNotExplainComponentNo,_FindSpecificationsError,
                    FindSequenceNumberError,FindStrainOrDispError,FindDescriptionError,
                    FindOtherBridgesWarnning,
                },
            };

            var config = XDocument.Load(@"Option.config");

            //Bug:修改配置文件后，读取结果不变
            foreach (var e in config.Elements("configuration").Descendants())
            {
                try
                {
                    controlClass.SelectListInt.Add(Convert.ToInt32(e.Value.ToString()));
                }
                catch (Exception)
                {
                    //TODO:增加读不出数据的异常处理
                    controlClass.SelectListInt.Add(0);
                }
            }

            var w = new ProgressBarWindow();
            w.Top = 0.4 * (App.ScreenHeight - w.Height);
            w.Left = 0.4 * (App.ScreenWidth - w.Width);

            var progressBarDataBinding = new ProgressBarDataBinding
            {
                V = 14,
            };
            w.progressBarNumberTextBlock.DataContext = progressBarDataBinding;
            w.progressBar.DataContext = progressBarDataBinding;

            var progressSleepTime = 100;    //进度条停顿时间

            var thread = new Thread(new ThreadStart(() =>
            {
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.Show(); });
                int invokeF = 0;    //已调用函数
                for (int i = 0; i < controlClass.SelectListInt.Count; i++)
                {
                    if (controlClass.SelectListInt[i] != 0)
                    {
                        controlClass.SelectListFunctionName[i]?.Invoke();
                        invokeF++;
                        progressBarDataBinding.V = invokeF * 100 / controlClass.SelectListInt.Sum();
                        //w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = invokeF * 100 / controlClass.SelectListInt.Sum(); });
                        //w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = invokeF*100/ controlClass.SelectListInt.Sum(); });
                        //w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBarNumber.Text = (invokeF * 100 / controlClass.SelectListInt.Sum()).ToString(); });
                        Thread.Sleep(progressSleepTime);
                    }

                }
                _GenerateResultReport();
                _doc.Save("标出错误或警告的报告.doc");
                Thread.Sleep(progressSleepTime);
                MessageBox.Show("已成功校核");
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.Close(); });
            }));
            thread.Start();

        }
    }
}
