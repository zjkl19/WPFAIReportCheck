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
        public delegate void CheckFunctionHandler();
        public event CheckFunctionHandler UpgradeProgressBar;

        private delegate void MyDelegate();
        public void CheckReport()
        {

            var fs = new List<MyDelegate>();
            fs.Add(this.FindDescriptionError);

            //举例：
            //SelectListInt =[1,0]
            //SelectListFunctionName=[FindUnitError,FindNotExplainComponentNo]
            //表示算法只算FindUnitError,不算FindNotExplainComponentNo
            var controlClass = new
            {
                SelectListInt = new List<int>(),
                SelectListFunctionName = new List<MyDelegate> {
                    _FindUnitError,_FindNotExplainComponentNo,_FindSpecificationsError,
                    FindSequenceNumberError,FindStrainOrDispError,FindDescriptionError,
                    FindOtherBridgesWarnning,
                },
            };

            
            var config = XDocument.Load(@"Option.config");
            var op = Convert.ToInt32(config.Element("configuration").Element("FindUnitError").Value);

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

            //foreach (XElement e in config.Elements("configuration"))
            //{
            //    try
            //    {
            //        controlClass.SelectListInt.Add(Convert.ToInt32(e.Value));
            //    }
            //    catch (Exception)
            //    {
            //        //TODO:增加读不出数据的异常处理
            //        controlClass.SelectListInt.Add(0);
            //    }   
            //}

            var w = new ProgressBarWindow();
            w.Top = 0.4 * (App.ScreenHeight - w.Height);
            w.Left = 0.4 * (App.ScreenWidth - w.Width);

            var progressSleepTime = 50;

            //CheckFunction cf;
            //cf = _FindUnitError;
            //cf += _FindNotExplainComponentNo;
            //cf += _FindSpecificationsError;
            UpgradeProgressBar += _FindUnitError;


            var thread = new Thread(new ThreadStart(() =>
            {
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.Show(); });
                int invokeF=0;    //已调用函数
                for(int i=0;i<controlClass.SelectListInt.Count;i++)
                {
                    if(controlClass.SelectListInt[i]!=0)
                    {
                        controlClass.SelectListFunctionName[i]?.Invoke();
                        invokeF++;
                        w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = invokeF*100/ controlClass.SelectListInt.Sum(); });
                        Thread.Sleep(progressSleepTime);
                    }
                    
                }
                //_FindUnitError();
                //w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 10; });
                //Thread.Sleep(progressSleepTime);
                //_FindNotExplainComponentNo();
                //w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 20; });
                //Thread.Sleep(progressSleepTime);
                //_FindSpecificationsError();
                //w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 30; });
                //Thread.Sleep(progressSleepTime);
                //FindSequenceNumberError();
                //w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 40; });
                //Thread.Sleep(progressSleepTime);
                //FindStrainOrDispError();
                //w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 50; });
                //Thread.Sleep(progressSleepTime);
                //FindDescriptionError();
                //w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 60; });
                //Thread.Sleep(progressSleepTime);
                //FindOtherBridgesWarnning();
                //w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.progressBar.Value = 70; });
                //Thread.Sleep(progressSleepTime);

                _GenerateResultReport();
                _doc.Save("标出错误或警告的报告.doc");
                Thread.Sleep(100);
                MessageBox.Show("已成功校核");
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.Close(); });
            }));
            thread.Start();



        }
    }
}
