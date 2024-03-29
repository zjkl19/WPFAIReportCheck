﻿using Aspose.Words;
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

        private string content = string.Empty;
        public int V { get
            { return v; } set
            {
                v = value;

                OnPropertyChanged("V");
            }
        }

        public string Content
        {
            get
            { return content; }
            set
            {
                content = value;

                OnPropertyChanged("Content");
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
        public delegate void MethodDelegate();
        public void CheckReport()
        {

            //举例：
            //SelectListInt =[1,0]
            //SelectListFunctionName=[FindUnitError,FindNotExplainComponentNo]
            //表示算法只算FindUnitError,不算FindNotExplainComponentNo
            var controlClass = RegisterMethod();

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

                try
                {
                    controlClass.SelectListContent.Add(e.Attribute("Content").Value.ToString());
                }
                catch (Exception)
                {
                    //TODO:增加读不出数据的异常处理
                    controlClass.SelectListContent.Add("方法名称读取失败");
                }
            }

            var w = new ProgressBarWindow();
            w.Top = 0.4 * (App.ScreenHeight - w.Height);
            w.Left = 0.4 * (App.ScreenWidth - w.Width);

            var progressBarDataBinding = new ProgressBarDataBinding
            {
                V = 0,
            };
            w.progressBarNumberTextBlock.DataContext = progressBarDataBinding;
            w.progressBar.DataContext = progressBarDataBinding;
            w.progressBarContentTextBlock.DataContext = progressBarDataBinding;

            var progressSleepTime = 500;    //进度条停顿时间

            var thread = new Thread(new ThreadStart(() =>
            {
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.Show(); });
                int invokeF = 0;    //已调用函数
                for (int i = 0; i < controlClass.SelectListInt.Count; i++)
                {
                    if (controlClass.SelectListInt[i] != 0)
                    {
                        
                        progressBarDataBinding.V = invokeF * 100 / controlClass.SelectListInt.Sum();
                        progressBarDataBinding.Content = $"正在校核：{controlClass.SelectListContent[i]}";
                        controlClass.SelectListFunctionName[i]?.Invoke();
                        invokeF++;
                        Thread.Sleep(progressSleepTime);
                    }

                }
                progressBarDataBinding.V = 100;
                progressBarDataBinding.Content = $"正在校核：正在完成中";
                GenerateResultReport();
                _doc.Save("标出错误或警告的报告.doc");
                Thread.Sleep(progressSleepTime);
                MessageBox.Show($"校核已完成！共校核出：" +
                    $"\r错误{reportError.Count}个" +
                    $"\r警告{reportWarnning.Count}个" +
                    $"\r信息{reportInfo.Count}条");
                w.progressBar.Dispatcher.BeginInvoke((ThreadStart)delegate { w.Close(); });
            }));
            thread.Start();

        }
    }   
}
