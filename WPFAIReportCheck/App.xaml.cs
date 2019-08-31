using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace WPFAIReportCheck
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        //参考https://www.cnblogs.com/Gildor/archive/2010/06/29/1767156.html
        public static double ScreenWidth = SystemParameters.PrimaryScreenWidth;
        public static double ScreenHeight = SystemParameters.PrimaryScreenHeight;
        //public static ILogger log;

        public App()
        {
            Startup += new StartupEventHandler(App_Startup);
        }

        //TODO:增加必要的全局变量和方法
        void App_Startup(object sender, StartupEventArgs e)
        {
        }
    }
}
